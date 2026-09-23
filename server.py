"""Word ⇄ PDF 手机网页版 —— 局域网服务端。

手机连同一 Wi-Fi，浏览器打开提示的网址即可上传转换、下载结果。
转换仍在本机完成：Word→PDF 走本机 Microsoft Word，输出质量与桌面版完全一致。

用法：
    env\\Scripts\\python.exe server.py
    env\\Scripts\\python.exe server.py --port 8765 --no-auth    # 关掉访问口令

设计要点：
* 纯标准库实现（http.server），不引入新依赖，现有 env 直接可跑。
* 上传用「原始二进制流」而非 multipart，避免解析开销与兼容问题。
* 转换在单一后台线程串行执行 —— Office COM 对多线程很敏感，串行最稳。
* 结果文件带随机令牌，只有发起转换的设备能下载。
"""

from __future__ import annotations

import argparse
import html
import io
import json
import os
import queue
import secrets
import shutil
import socket
import sys
import threading
import time
import traceback
import urllib.parse
from dataclasses import dataclass, field
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


def _bootstrap_env() -> None:
    """确保使用程序自带的 env（与 app.py 相同的保护逻辑）。"""
    if os.environ.get("WPC_REEXEC") == "1":
        return
    env_dir = os.path.join(HERE, "env")
    try:
        if os.path.normcase(os.path.abspath(sys.prefix)) == os.path.normcase(env_dir):
            return
    except Exception:  # noqa: BLE001
        pass
    venv_py = os.path.join(env_dir, "Scripts", "python.exe")
    if not os.path.isfile(venv_py):
        return
    system32 = os.path.join(env_dir, "Lib", "site-packages", "pywin32_system32")
    if not os.path.isdir(system32):
        return
    os.environ["WPC_REEXEC"] = "1"
    try:
        os.execv(venv_py, [venv_py, os.path.abspath(__file__), *sys.argv[1:]])
    except Exception:  # noqa: BLE001
        os.environ.pop("WPC_REEXEC", None)


_bootstrap_env()

import converter  # noqa: E402

APP_NAME = "Word ⇄ PDF 手机互转"
MAX_UPLOAD_BYTES = 200 * 1024 * 1024      # 单文件 200 MB
RESULT_TTL_SECONDS = 2 * 60 * 60          # 结果保留 2 小时
MAX_JOBS_KEPT = 200


# ---------------------------------------------------------------- 任务管理


@dataclass
class Job:
    """一次转换任务。"""

    job_id: str
    token: str
    filename: str
    direction: str
    src: str
    dst: str = ""
    status: str = "排队中"          # 排队中 / 转换中 / 完成 / 失败 / 已取消
    progress: str = ""
    error: str = ""
    created: float = field(default_factory=time.time)
    finished: float = 0.0

    def public(self) -> dict:
        data = {
            "id": self.job_id,
            "filename": self.filename,
            "direction": self.direction,
            "status": self.status,
            "progress": self.progress,
            "error": self.error,
        }
        if self.status == "完成" and self.dst and os.path.isfile(self.dst):
            data["size"] = os.path.getsize(self.dst)
            data["download"] = f"/api/download?id={self.job_id}&token={self.token}"
            data["result_name"] = os.path.basename(self.dst)
        return data


class JobManager:
    """串行执行转换任务的后台管理器。"""

    def __init__(self, engine: str = "auto") -> None:
        self.jobs: dict[str, Job] = {}
        self.lock = threading.Lock()
        self.queue: queue.Queue[str] = queue.Queue()
        # Word→PDF 使用的引擎：auto / office / libreoffice
        self.engine = engine
        self.workdir = os.path.join(converter.temp_root(), "web")
        os.makedirs(self.workdir, exist_ok=True)
        self.worker = threading.Thread(target=self._run, daemon=True)
        self.worker.start()

    # -------------------------------------------------- 对外接口

    def submit(self, filename: str, direction: str, token: str,
               stream, length: int) -> Job:
        job_id = secrets.token_hex(8)
        safe_name = _safe_filename(filename)
        job_dir = os.path.join(self.workdir, job_id)
        os.makedirs(job_dir, exist_ok=True)
        src = os.path.join(job_dir, safe_name)

        written = 0
        with open(src, "wb") as fh:
            remaining = length
            while remaining > 0:
                chunk = stream.read(min(262144, remaining))
                if not chunk:
                    break
                fh.write(chunk)
                written += len(chunk)
                remaining -= len(chunk)
        if written == 0:
            raise ValueError("上传内容为空")

        job = Job(job_id=job_id, token=token, filename=safe_name,
                  direction=direction, src=src)
        dst_name = os.path.splitext(safe_name)[0] + (
            ".pdf" if direction == "word2pdf" else ".docx")
        job.dst = os.path.join(job_dir, dst_name)
        with self.lock:
            self.jobs[job_id] = job
            self._prune_locked()
        self.queue.put(job_id)
        return job

    def get(self, job_id: str) -> Job | None:
        with self.lock:
            return self.jobs.get(job_id)

    def recent(self, limit: int = 20) -> list[dict]:
        """最近的任务（按时间倒序），供手机端重新下载。"""
        with self.lock:
            jobs = sorted(self.jobs.values(), key=lambda j: j.created, reverse=True)
            return [j.public() for j in jobs[:limit]]

    def stats(self) -> dict:
        with self.lock:
            pending = sum(1 for j in self.jobs.values() if j.status in ("排队中", "转换中"))
        return {
            "queued": pending,
            "engine": converter.describe_engine(),
            "engine_mode": self.engine,
            # 按实际使用的引擎模式判断，而不是「本机有没有某个引擎」，
            # 否则强制 libreoffice 但本机只有 Word 时会误报可用
            "can_word2pdf": self.can_word2pdf(),
        }

    def can_word2pdf(self) -> bool:
        """当前引擎模式下 Word→PDF 是否真的可用。"""
        if self.engine == "office":
            return bool(converter.word_engine_available()
                        or converter.com_engine_progid())
        if self.engine == "libreoffice":
            return bool(converter.find_soffice())
        return bool(converter.word_engine_available()
                    or converter.com_engine_progid()
                    or converter.find_soffice())

    # -------------------------------------------------- 内部

    def _prune_locked(self) -> None:
        """清理过期任务与它们的临时文件。"""
        now = time.time()
        dead = []
        for job_id, job in self.jobs.items():
            expired = job.finished and (now - job.finished) > RESULT_TTL_SECONDS
            if expired:
                dead.append(job_id)
        if len(self.jobs) - len(dead) > MAX_JOBS_KEPT:
            ordered = sorted(self.jobs.values(), key=lambda j: j.created)
            for job in ordered:
                if len(self.jobs) - len(dead) <= MAX_JOBS_KEPT:
                    break
                if job.job_id not in dead:
                    dead.append(job.job_id)
        for job_id in dead:
            job = self.jobs.pop(job_id, None)
            if job:
                shutil.rmtree(os.path.dirname(job.src), ignore_errors=True)

    def _run(self) -> None:
        while True:
            job_id = self.queue.get()
            job = self.get(job_id)
            if job is None:
                continue
            job.status = "转换中"
            job.progress = "开始处理…"
            try:
                outcomes = converter.run_jobs(
                    [converter.Job(src=job.src, dst=job.dst, direction=job.direction)],
                    progress=lambda text, j=job: setattr(j, "progress", text),
                    engine=self.engine,
                )
                _, error = outcomes[0]
                if error:
                    job.status = "已取消" if error == "已取消" else "失败"
                    job.error = error
                else:
                    job.status = "完成"
                    job.progress = "转换完成"
            except Exception as exc:  # noqa: BLE001
                job.status = "失败"
                job.error = f"{type(exc).__name__}: {exc}"
                job.progress = ""
                print("[错误] 转换任务异常：\n" + traceback.format_exc())
            finally:
                job.finished = time.time()
                # 源文件只在成功后才删（失败时保留，便于排查）
                try:
                    if job.status == "完成":
                        if os.path.isfile(job.src):
                            os.remove(job.src)
                except Exception:  # noqa: BLE001
                    pass


def _safe_filename(name: str) -> str:
    """去掉路径成分与非法字符，防止目录穿越。"""
    name = os.path.basename(name.replace("\\", "/")).strip()
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, "_")
    if not name or name in (".", ".."):
        name = "upload"
    return name[:120]


# ---------------------------------------------------------------- HTTP 服务


class Server(ThreadingHTTPServer):
    """带守护线程的 HTTP 服务，避免手机端断连时挂住进程退出。"""

    daemon_threads = True
    allow_reuse_address = True


class Handler(BaseHTTPRequestHandler):
    server_version = "WordPdfMobile/1.0"
    protocol_version = "HTTP/1.1"

    manager: JobManager
    password: str | None = None
    lan_url: str = ""

    # ------------------------------------------------ 工具

    def log_message(self, fmt: str, *args) -> None:
        """精简访问日志。"""
        try:
            print(f"  [{time.strftime('%H:%M:%S')}] {self.address_string()} "
                  f"{fmt % args}")
        except Exception:  # noqa: BLE001
            pass

    def _send(self, code: int, body: bytes, content_type: str,
              extra: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass   # 手机上取消下载很正常

    def _json(self, code: int, payload: dict) -> None:
        self._send(code, json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _query(self) -> dict:
        parsed = urllib.parse.urlparse(self.path)
        return {k: v[0] for k, v in urllib.parse.parse_qs(parsed.query).items()}

    def _path(self) -> str:
        return urllib.parse.urlparse(self.path).path

    def _check_auth(self, query: dict) -> bool:
        if not Handler.password:
            return True
        supplied = (self.headers.get("X-Password") or query.get("pw") or "")
        return secrets.compare_digest(supplied, Handler.password)

    # ------------------------------------------------ 路由

    def do_HEAD(self) -> None:  # noqa: N802
        """部分浏览器会先发 HEAD 探测，简单应答即可。"""
        if self._path() in ("/", "/index.html"):
            body = PAGE_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return
        self.send_response(404)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        path = self._path()
        query = self._query()
        if path in ("/", "/index.html"):
            self._send(200, PAGE_HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        # PWA 资源（无需口令，浏览器要能直接取到才能安装）
        if path == "/manifest.webmanifest":
            self._send(200, json.dumps(MANIFEST, ensure_ascii=False).encode("utf-8"),
                       "application/manifest+json; charset=utf-8")
            return
        if path == "/icon.svg":
            self._send(200, ICON_SVG.encode("utf-8"), "image/svg+xml; charset=utf-8")
            return
        if path == "/sw.js":
            self._send(200, SERVICE_WORKER.encode("utf-8"),
                       "application/javascript; charset=utf-8")
            return
        if path == "/favicon.ico":
            self._send(200, ICON_SVG.encode("utf-8"), "image/svg+xml; charset=utf-8")
            return
        if path == "/api/info":
            info = self.manager.stats()
            info["need_password"] = bool(Handler.password)
            info["url"] = Handler.lan_url
            self._json(200, info)
            return
        if not self._check_auth(query):
            self._json(HTTPStatus.UNAUTHORIZED, {"error": "访问口令不正确"})
            return
        if path == "/api/status":
            job = self.manager.get(query.get("id", ""))
            if job is None:
                self._json(HTTPStatus.NOT_FOUND, {"error": "任务不存在或已过期"})
                return
            self._json(200, job.public())
            return
        if path == "/api/download":
            self._download(query)
            return
        if path == "/api/recent":
            self._json(200, {"jobs": self.manager.recent()})
            return
        self._json(HTTPStatus.NOT_FOUND, {"error": "未知接口"})

    def do_POST(self) -> None:  # noqa: N802
        path = self._path()
        query = self._query()
        if not self._check_auth(query):
            self._json(HTTPStatus.UNAUTHORIZED, {"error": "访问口令不正确"})
            return
        if path != "/api/convert":
            self._json(HTTPStatus.NOT_FOUND, {"error": "未知接口"})
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            self._json(HTTPStatus.BAD_REQUEST, {"error": "没有收到文件内容"})
            return
        if length > MAX_UPLOAD_BYTES:
            self._json(HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                       {"error": f"文件超过 {MAX_UPLOAD_BYTES // 1024 // 1024} MB 上限"})
            return

        filename = urllib.parse.unquote(self.headers.get("X-Filename") or "upload")
        direction = (self.headers.get("X-Direction") or "").strip()
        if direction not in ("word2pdf", "pdf2word"):
            kind = converter.file_kind(filename)
            if kind == "word":
                direction = "word2pdf"
            elif kind == "pdf":
                direction = "pdf2word"
            else:
                self._json(HTTPStatus.BAD_REQUEST,
                           {"error": "无法识别的文件类型，请选择 .docx/.doc 或 .pdf"})
                return
        # 方向与扩展名不一致时以扩展名为准，避免误转换
        kind = converter.file_kind(filename)
        if kind == "unknown":
            self._json(HTTPStatus.BAD_REQUEST, {"error": "不支持的文件类型"})
            return
        if direction == "word2pdf" and kind != "word":
            self._json(HTTPStatus.BAD_REQUEST, {"error": "这是 PDF 文件，请选择 PDF → Word"})
            return
        if direction == "pdf2word" and kind != "pdf":
            self._json(HTTPStatus.BAD_REQUEST, {"error": "这不是 PDF 文件，请选择 Word → PDF"})
            return

        try:
            job = self.manager.submit(filename, direction,
                                      query.get("token", ""), self.rfile, length)
        except Exception as exc:  # noqa: BLE001
            self._json(HTTPStatus.INTERNAL_SERVER_ERROR,
                       {"error": f"接收文件失败：{exc}"})
            return
        self._json(200, job.public())

    def _download(self, query: dict) -> None:
        job = self.manager.get(query.get("id", ""))
        if job is None:
            self._json(HTTPStatus.NOT_FOUND, {"error": "任务不存在或已过期"})
            return
        if query.get("token") != job.token:
            self._json(HTTPStatus.FORBIDDEN, {"error": "下载链接无效"})
            return
        if job.status != "完成" or not os.path.isfile(job.dst):
            self._json(HTTPStatus.CONFLICT, {"error": "文件还没转换好"})
            return
        size = os.path.getsize(job.dst)
        # 中文文件名需要 RFC 5987 编码
        quoted = urllib.parse.quote(os.path.basename(job.dst))
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/pdf" if job.dst.lower().endswith(".pdf")
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header("Content-Length", str(size))
        self.send_header("Content-Disposition",
                         f"attachment; filename*=UTF-8''{quoted}")
        self.end_headers()
        try:
            with open(job.dst, "rb") as fh:
                shutil.copyfileobj(fh, self.wfile, 262144)
        except (BrokenPipeError, ConnectionResetError):
            pass


# ---------------------------------------------------------------- 前端页面

MANIFEST = {
    "name": "Word ⇄ PDF 互转",
    "short_name": "Word⇄PDF",
    "description": "在手机上把 Word 与 PDF 互相转换（转换在电脑本机完成）",
    "lang": "zh-CN",
    "start_url": "./",
    "scope": "./",
    "display": "standalone",
    "orientation": "portrait",
    "background_color": "#0d1117",
    "theme_color": "#1f6feb",
    "icons": [{"src": "icon.svg", "sizes": "any",
               "type": "image/svg+xml", "purpose": "any maskable"}],
}

ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
<rect width="512" height="512" rx="112" fill="#1f6feb"/>
<path d="M136 96h150l90 90v230a24 24 0 0 1-24 24H136a24 24 0 0 1-24-24V120a24 24 0 0 1 24-24z" fill="#fff" opacity=".95"/>
<path d="M286 96l90 90h-66a24 24 0 0 1-24-24z" fill="#fff" opacity=".6"/>
<path d="M176 238h116l-22-24h34l40 44-40 44h-34l22-24H176z" fill="#1f6feb"/>
<path d="M336 330H220l22 24h-34l-40-44 40-44h34l-22 24h116z" fill="#1f6feb"/>
<text x="256" y="440" font-family="Segoe UI,Microsoft YaHei,sans-serif" font-size="74"
 font-weight="700" fill="#fff" text-anchor="middle">PDF</text>
</svg>
"""

# 极简 Service Worker：仅用于满足浏览器「可安装」的条件，不做离线缓存
# （转换必须联网，缓存页面反而容易让人误以为离线也能用）
SERVICE_WORKER = """self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => self.clients.claim());
self.addEventListener('fetch', e => {});
"""

PAGE_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#1f6feb">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Word⇄PDF">
<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="icon.svg">
<title>Word ⇄ PDF 互转</title>
<style>
:root{
  --bg:#0d1117; --card:#161b22; --line:#2a3240; --fg:#e6edf3; --dim:#8b949e;
  --accent:#2f81f7; --ok:#3fb950; --err:#f85149; --warn:#d29922;
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;padding:0}
body{
  background:var(--bg); color:var(--fg); min-height:100vh;
  font:16px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",
       "Hiragino Sans GB","Microsoft YaHei",sans-serif;
  padding:env(safe-area-inset-top) env(safe-area-inset-right)
          env(safe-area-inset-bottom) env(safe-area-inset-left);
}
.wrap{max-width:680px;margin:0 auto;padding:20px 16px 40px}
header{text-align:center;padding:12px 0 22px}
h1{font-size:21px;margin:0 0 6px;font-weight:600}
.sub{color:var(--dim);font-size:13px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;
      padding:16px;margin-bottom:14px}
.seg{display:flex;gap:8px;margin-bottom:4px}
.seg label{flex:1;text-align:center;padding:12px 6px;border:1px solid var(--line);
  border-radius:11px;font-size:15px;color:var(--dim);cursor:pointer;transition:.15s}
.seg input{display:none}
.seg input:checked+span{color:#fff}
.seg label:has(input:checked){background:var(--accent);border-color:var(--accent);color:#fff}
.drop{
  border:2px dashed var(--line);border-radius:14px;padding:34px 16px;text-align:center;
  color:var(--dim);cursor:pointer;transition:.18s;background:#11161d
}
.drop:active,.drop.hot{border-color:var(--accent);background:#132033;color:var(--fg)}
.drop .big{font-size:34px;line-height:1;margin-bottom:10px;display:block}
.hint{font-size:12.5px;color:var(--dim);margin-top:8px}
.file{
  display:flex;align-items:center;gap:10px;padding:12px;border:1px solid var(--line);
  border-radius:11px;margin-bottom:10px;background:#11161d;word-break:break-all
}
.file .nm{flex:1;font-size:14px}
.file .sz{color:var(--dim);font-size:12.5px;white-space:nowrap}
button{
  width:100%;padding:15px;border:0;border-radius:12px;font-size:16.5px;font-weight:600;
  background:var(--accent);color:#fff;cursor:pointer;transition:.15s
}
button:disabled{opacity:.45;cursor:not-allowed}
button.ghost{background:transparent;border:1px solid var(--line);color:var(--fg);
  font-weight:400;font-size:15px;margin-top:10px}
button.ok{background:var(--ok)}
.bar{height:7px;background:#0b0f14;border-radius:99px;overflow:hidden;margin:12px 0 9px}
.bar i{display:block;height:100%;width:0;background:var(--accent);transition:width .3s}
.bar.indet i{width:38%;animation:slide 1.15s ease-in-out infinite}
@keyframes slide{0%{margin-left:-40%}100%{margin-left:104%}}
.st{font-size:13.5px;color:var(--dim);min-height:20px;word-break:break-word}
.st.err{color:var(--err)} .st.ok{color:var(--ok)}
.note{font-size:12.5px;color:var(--dim);margin-top:12px;text-align:center}
.pill{display:inline-block;padding:3px 9px;border-radius:99px;background:#1c2431;
  color:var(--dim);font-size:12px;margin-top:4px}
.hidden{display:none}
.card.warn{border-color:#6b4a12;background:#241c0c;color:#e3b341}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Word ⇄ PDF 互转</h1>
    <div class="sub" id="engine">正在检测转换引擎…</div>
  </header>

  <div class="card warn hidden" id="warncard">
    <div id="warntext"></div>
  </div>

  <div class="card" id="pwcard" style="display:none">
    <div style="font-size:14px;margin-bottom:10px">请输入访问口令（在电脑上运行的程序窗口里）</div>
    <input id="pw" type="password" inputmode="text" autocomplete="off"
      style="width:100%;padding:13px;border-radius:10px;border:1px solid var(--line);
             background:#0b0f14;color:var(--fg);font-size:16px">
    <button class="ghost" id="pwsave">确定</button>
  </div>

  <div class="card">
    <div class="seg">
      <label><input type="radio" name="dir" value="word2pdf" checked><span>Word → PDF</span></label>
      <label><input type="radio" name="dir" value="pdf2word"><span>PDF → Word</span></label>
    </div>
    <div class="hint" style="text-align:center;margin-bottom:14px" id="dirhint">
      选择或拍照扫描的 <b>.docx / .doc</b> 文件
    </div>

    <div class="drop" id="drop">
      <span class="big">📄</span>
      <div id="droptext">点这里选择文件</div>
    </div>
    <input type="file" id="file" class="hidden">

    <div id="fileselect" style="margin-top:14px"></div>
    <button id="go" disabled>开始转换</button>
    <button id="reset" class="ghost hidden">转换另一个文件</button>

    <div class="bar hidden" id="barwrap"><i id="bar"></i></div>
    <div class="st" id="status"></div>
    <div class="note" id="note"></div>
  </div>

  <div class="card">
    <button class="ghost" id="histbtn" style="margin-top:0">查看最近转换记录</button>
    <div id="histlist" style="margin-top:12px"></div>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
const state = {file:null, job:null, token:null, timer:null, password:'', canWord2Pdf:true};

function fmtSize(n){
  if(n < 1024) return n + ' B';
  if(n < 1048576) return (n/1024).toFixed(1) + ' KB';
  return (n/1048576).toFixed(2) + ' MB';
}
function setStatus(text, cls){
  const el = $('status');
  el.textContent = text || '';
  el.className = 'st' + (cls ? ' ' + cls : '');
}
function setBar(pct, indeterminate){
  const wrap = $('barwrap'), bar = $('bar');
  if(pct === null){ wrap.classList.add('hidden'); return; }
  wrap.classList.remove('hidden');
  if(indeterminate){ wrap.classList.add('indet'); bar.style.width = ''; }
  else { wrap.classList.remove('indet'); bar.style.width = pct + '%'; }
}

// ---------- 引擎信息 ----------
async function loadInfo(){
  try{
    const r = await fetch('/api/info');
    const d = await r.json();
    $('engine').textContent = d.engine || '';
    state.canWord2Pdf = d.can_word2pdf !== false;
    applyEngineState();
    if(d.need_password){
      if(!state.password) state.password = localStorage.getItem('wpc_pw') || '';
      $('pwcard').style.display = state.password ? 'none' : 'block';
    }
    return d;
  }catch(e){ $('engine').textContent = '无法连接电脑服务'; }
}

// 没有 Word 引擎时，禁用「Word → PDF」并说明原因，避免用户白试
function applyEngineState(){
  const radio = document.querySelector('input[name=dir][value=word2pdf]');
  const label = radio ? radio.closest('label') : null;
  const warn = $('warncard');
  if(state.canWord2Pdf){
    if(label){ label.style.opacity = ''; label.style.pointerEvents = ''; }
    warn.classList.add('hidden');
    if(radio && radio.checked === false && !state.file) { /* 保持用户选择 */ }
    return;
  }
  if(label){ label.style.opacity = '.4'; label.style.pointerEvents = 'none'; }
  warn.classList.remove('hidden');
  $('warntext').innerHTML =
    '⚠ 这台电脑没有可用的 Word 引擎（未装 Microsoft Word 或 WPS），' +
    '<b>Word → PDF 暂时不可用</b>。<br>PDF → Word 仍可正常使用。';
  const pdfRadio = document.querySelector('input[name=dir][value=pdf2word]');
  if(pdfRadio && !pdfRadio.checked){
    pdfRadio.checked = true;
    pdfRadio.dispatchEvent(new Event('change'));
  }
}
$('pwsave').onclick = () => {
  state.password = $('pw').value.trim();
  localStorage.setItem('wpc_pw', state.password);
  $('pwcard').style.display = 'none';
  loadInfo();
};

// ---------- 方向切换 ----------
document.querySelectorAll('input[name=dir]').forEach(r => {
  r.onchange = () => {
    const w = r.value === 'word2pdf';
    $('dirhint').innerHTML = w
      ? '选择或拍照扫描的 <b>.docx / .doc</b> 文件'
      : '选择 <b>.pdf</b> 文件，转换为可编辑的 Word';
    $('file').accept = w ? '.docx,.doc,.docm,.rtf,.odt,.txt' : '.pdf';
    state.file = null; renderFile();
  };
});
function currentDir(){ return document.querySelector('input[name=dir]:checked').value; }

// ---------- 选择文件 ----------
const accept = () => currentDir() === 'word2pdf'
  ? '.docx,.doc,.docm,.rtf,.odt,.txt' : '.pdf';
$('file').accept = accept();

$('drop').onclick = () => $('file').click();
$('file').onchange = e => { pick(e.target.files[0]); };

['dragenter','dragover'].forEach(ev => $('drop').addEventListener(ev, e => {
  e.preventDefault(); $('drop').classList.add('hot');
}));
['dragleave','drop'].forEach(ev => $('drop').addEventListener(ev, e => {
  e.preventDefault(); $('drop').classList.remove('hot');
}));
$('drop').addEventListener('drop', e => {
  if(e.dataTransfer.files.length) pick(e.dataTransfer.files[0]);
});

function pick(f){
  if(!f) return;
  const isPdf = /\.pdf$/i.test(f.name);
  const dir = currentDir();
  if(dir === 'word2pdf' && !state.canWord2Pdf){
    setStatus('这台电脑没有可用的 Word 引擎，请改用「PDF → Word」', 'err');
    return;
  }
  if(dir === 'word2pdf' && isPdf){
    setStatus('这是 PDF 文件，请在上方选择「PDF → Word」', 'err');
    return;
  }
  if(dir === 'pdf2word' && !isPdf){
    setStatus('这不是 PDF 文件，请在上方选择「Word → PDF」', 'err');
    return;
  }
  state.file = f;
  setStatus('');
  renderFile();
}
function renderFile(){
  const box = $('fileselect');
  if(!state.file){ box.innerHTML = ''; $('go').disabled = true; return; }
  box.innerHTML = `<div class="file"><span>📎</span>
     <span class="nm">${escapeHtml(state.file.name)}</span>
     <span class="sz">${fmtSize(state.file.size)}</span></div>`;
  $('go').disabled = false;
}
function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c => (
    {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ---------- 上传 ----------
$('go').onclick = () => {
  if(!state.file) return;
  const dir = currentDir();
  const token = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const qs = new URLSearchParams({token});
  if(state.password) qs.set('pw', state.password);

  $('go').disabled = true;
  $('go').textContent = '正在上传…';
  setBar(0, false);
  setStatus('上传中…');

  const xhr = new XMLHttpRequest();
  xhr.open('POST', '/api/convert?' + qs.toString());
  xhr.setRequestHeader('X-Filename', encodeURIComponent(state.file.name));
  xhr.setRequestHeader('X-Direction', dir);
  if(state.password) xhr.setRequestHeader('X-Password', state.password);

  xhr.upload.onprogress = e => {
    if(e.lengthComputable){
      const pct = Math.round(e.loaded / e.total * 100);
      setBar(pct * 0.35, false);
      setStatus(`上传中… ${pct}%`);
    }
  };
  xhr.onload = () => {
    let d = {};
    try{ d = JSON.parse(xhr.responseText); }catch(e){}
    if(xhr.status !== 200){
      fail(d.error || ('上传失败（HTTP ' + xhr.status + '）'));
      if(xhr.status === 401){ $('pwcard').style.display = 'block'; }
      return;
    }
    state.job = d.id; state.token = token;
    setBar(38, false);
    setStatus('已排队，等待转换…');
    poll();
  };
  xhr.onerror = () => fail('网络中断，请确认手机与电脑在同一 Wi-Fi');
  xhr.send(state.file);
};

function fail(msg){
  setBar(null);
  setStatus(msg, 'err');
  $('go').disabled = false;
  $('go').textContent = '开始转换';
}

// ---------- 轮询进度 ----------
function poll(){
  clearTimeout(state.timer);
  const qs = new URLSearchParams({id: state.job, token: state.token});
  if(state.password) qs.set('pw', state.password);
  state.timer = setTimeout(async () => {
    try{
      const r = await fetch('/api/status?' + qs.toString());
      const d = await r.json();
      if(r.status !== 200){ fail(d.error || '任务查询失败'); return; }
      applyStatus(d);
    }catch(e){
      state.timer = setTimeout(poll, 1500);   // 手机锁屏后可能瞬时断网，重试
    }
  }, 900);
}

function applyStatus(d){
  if(d.status === '排队中'){
    $('go').textContent = '排队中…';
    setBar(38, true);
    setStatus(d.progress || '前面还有任务，请稍候…');
    poll();
    return;
  }
  if(d.status === '转换中'){
    $('go').textContent = '转换中…';
    setBar(null); setBar(42, true);
    setStatus(d.progress || '正在转换…');
    poll();
    return;
  }
  if(d.status === '完成'){
    setBar(100, false);
    setStatus('转换完成，正在下载…', 'ok');
    $('go').textContent = '已下载';
    $('reset').classList.remove('hidden');
    $('note').innerHTML = '若没有自动开始下载，'
      + `<a style="color:#2f81f7" href="${d.download}">点这里手动下载</a>`;
    // 触发下载
    const a = document.createElement('a');
    a.href = d.download;
    a.download = d.result_name || '';
    document.body.appendChild(a); a.click(); a.remove();
    return;
  }
  // 失败 / 已取消
  fail(d.error || d.status);
  $('reset').classList.remove('hidden');
}

$('reset').onclick = () => {
  clearTimeout(state.timer);
  state.file = null; state.job = null;
  $('file').value = '';
  renderFile();
  $('go').textContent = '开始转换';
  $('go').disabled = true;
  $('reset').classList.add('hidden');
  $('note').textContent = '';
  setBar(null);
  setStatus('');
};

// ---------- 最近转换记录 ----------
$('histbtn').onclick = async () => {
  const box = $('histlist');
  if(box.dataset.open === '1'){
    box.innerHTML = ''; box.dataset.open = '0';
    $('histbtn').textContent = '查看最近转换记录';
    return;
  }
  $('histbtn').textContent = '正在加载…';
  const qs = new URLSearchParams();
  if(state.password) qs.set('pw', state.password);
  try{
    const r = await fetch('/api/recent?' + qs.toString());
    if(r.status === 401){ $('pwcard').style.display = 'block'; throw new Error('需要口令'); }
    const d = await r.json();
    const done = (d.jobs || []).filter(j => j.status === '完成' && j.download);
    if(!done.length){
      box.innerHTML = '<div class="hint">还没有可下载的记录。转换完成后会显示在这里（保留 2 小时）。</div>';
    } else {
      box.innerHTML = done.map(j => `
        <div class="file">
          <span>${j.direction === 'word2pdf' ? '📄' : '📝'}</span>
          <span class="nm">${escapeHtml(j.result_name || j.filename)}
            <div class="hint" style="margin:0">${escapeHtml(j.filename)}</div>
          </span>
          <a style="color:#2f81f7;font-size:14px;white-space:nowrap"
             href="${j.download}${state.password ? '&pw=' + encodeURIComponent(state.password) : ''}"
             download>下载</a>
        </div>`).join('');
    }
    box.dataset.open = '1';
    $('histbtn').textContent = '收起记录';
  }catch(e){
    box.innerHTML = '<div class="hint st err">加载失败：' + escapeHtml(e.message) + '</div>';
    $('histbtn').textContent = '重试';
  }
};

loadInfo();
if('serviceWorker' in navigator){
  navigator.serviceWorker.register('sw.js').catch(()=>{});
}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------- 启动


def lan_ip() -> str:
    """尽力获取本机在局域网中的 IP。"""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(0.5)
        sock.connect(("8.8.8.8", 80))       # 不会真正发包
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except Exception:  # noqa: BLE001
        pass
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:  # noqa: BLE001
        return "127.0.0.1"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="server.py",
        description="Word ⇄ PDF 手机网页版（局域网服务端）",
    )
    parser.add_argument("--port", type=int, default=8765, help="监听端口（默认 8765）")
    parser.add_argument("--host", default="0.0.0.0",
                        help="监听地址（默认 0.0.0.0，供局域网访问）")
    parser.add_argument("--no-auth", action="store_true", help="关闭访问口令")
    parser.add_argument("--password", default=None, help="指定访问口令")
    parser.add_argument("--engine", choices=("auto", "office", "libreoffice"),
                        default="auto",
                        help="Word→PDF 引擎：auto（有 Word 就用）/ office（强制 Word 或 WPS）"
                             " / libreoffice（强制 LibreOffice，适合部署到无 Word 的服务器）")
    args = parser.parse_args(argv)

    converter.cleanup_temp()

    if args.no_auth:
        Handler.password = None
    else:
        Handler.password = args.password or "".join(
            secrets.choice("0123456789") for _ in range(6))

    ip = lan_ip()
    Handler.lan_url = f"http://{ip}:{args.port}/"
    manager = JobManager(engine=args.engine)
    Handler.manager = manager

    # 引擎警告：强制单引擎但该引擎不可用时，提前说清楚
    warning = ""
    if args.engine == "office" and not converter.word_engine_available() \
            and not converter.com_engine_progid():
        warning = "指定了 office 引擎，但本机没检测到 Word/WPS"
    elif args.engine == "libreoffice" and not converter.find_soffice():
        warning = "指定了 libreoffice 引擎，但本机没检测到 LibreOffice"
    elif args.engine == "auto" and not converter.word_engine_available() \
            and not converter.com_engine_progid() and not converter.find_soffice():
        warning = "本机没有任何可用引擎，Word→PDF 会失败"

    try:
        httpd = Server((args.host, args.port), Handler)
    except OSError as exc:
        print(f"\n[错误] 无法监听 {args.host}:{args.port} —— {exc}")
        print("      端口可能被占用，可换一个：--port 8899")
        return 1

    line = "=" * 56
    print(f"\n{line}")
    print(f"  {APP_NAME} —— 手机网页版")
    print(line)
    print(f"  转换引擎 : {converter.describe_engine()}")
    if args.engine != "auto":
        print(f"  引擎模式 : 强制 {args.engine}")
    if warning:
        print(f"  ⚠ 注意   : {warning}")
    print(f"  手机访问 : {Handler.lan_url}")
    if Handler.password:
        print(f"  访问口令 : {Handler.password}")
    else:
        print("  访问口令 : 已关闭（同网络内任何人都可用）")
    print(line)
    print("  手机与电脑需连同一个 Wi-Fi。")
    print("  保持本窗口开着，按 Ctrl+C 可停止服务。")
    print(f"{line}\n")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n正在停止服务…")
    finally:
        httpd.shutdown()
        shutil.rmtree(manager.workdir, ignore_errors=True)
        converter.cleanup_temp()
        print("服务已停止。")
    return 0


if __name__ == "__main__":
    try:
        code = main()
    finally:
        pass
    sys.exit(code)
