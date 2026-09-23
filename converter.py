"""Word <-> PDF 转换核心逻辑。

本模块不依赖图形界面，可被界面或命令行单独调用。

引擎说明
--------
* Word -> PDF：调用本机 Microsoft Word 的 COM 自动化接口导出，由 Word 自身排版，
  输出与「在 Word 中另存为 PDF」完全一致，文字可选可搜索。
  若本机没有 Word，则尝试 LibreOffice 命令行（soffice）作为备用引擎。
* PDF -> Word：用 PyMuPDF 解析版面（字号、粗体、分栏、表格、图片），
  再由 python-docx 生成真正可编辑的 .docx。

依赖（全部装在程序自带 env 目录内）：pywin32、PyMuPDF、python-docx。
"""

from __future__ import annotations

import io
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from typing import Callable, Sequence

# ---------------------------------------------------------------- 常量

WD_FORMAT_PDF = 17  # wdExportFormatPDF

WORD_EXTS = (".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt", ".txt", ".wps")
PDF_EXTS = (".pdf",)

ProgressCb = Callable[[str], None]
CancelCb = Callable[[], bool]

# 本文件所在目录（程序根目录）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 运行时产生的临时文件统一放在程序目录下的 .tmp，
# 避免依赖系统临时目录（某些受限环境下系统 TEMP 不可写）。
TEMP_DIRNAME = ".tmp"


def temp_root() -> str:
    """返回可用的临时目录，优先使用程序目录下的 .tmp。"""
    candidates = [os.path.join(BASE_DIR, TEMP_DIRNAME)]
    env_tmp = os.environ.get("TEMP") or os.environ.get("TMP")
    if env_tmp:
        candidates.append(os.path.join(env_tmp, "WordPdfConverter"))
    for candidate in candidates:
        try:
            os.makedirs(candidate, exist_ok=True)
            probe = os.path.join(candidate, ".probe")
            with open(probe, "w", encoding="utf-8") as fh:
                fh.write("ok")
            os.remove(probe)
            return candidate
        except Exception:  # noqa: BLE001
            continue
    return tempfile.gettempdir()


def cleanup_temp() -> None:
    """清理程序目录下遗留的临时文件（异常退出后可能残留）。"""
    target = os.path.join(BASE_DIR, TEMP_DIRNAME)
    if not os.path.isdir(target):
        return
    for name in os.listdir(target):
        full = os.path.join(target, name)
        try:
            if os.path.isdir(full):
                shutil.rmtree(full, ignore_errors=True)
            else:
                os.remove(full)
        except Exception:  # noqa: BLE001
            continue


class ConvertError(RuntimeError):
    """转换失败，消息面向最终用户。"""


class Cancelled(ConvertError):
    """用户主动取消。"""


def file_kind(path: str) -> str:
    """返回 'word' / 'pdf' / 'unknown'。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in PDF_EXTS:
        return "pdf"
    if ext in WORD_EXTS:
        return "word"
    return "unknown"


# ---------------------------------------------------------------- 引擎探测

# Microsoft Word 的 COM ProgID
WORD_PROGIDS = ("Word.Application",)
# WPS Office 的 Word 兼容 ProgID（中文系统常见）
WPS_PROGIDS = ("KWPS.Application", "WPS.Application")

_WINWORD_GUESSES = (
    r"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office14\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE",
)

_WPS_GUESSES = (
    r"C:\Program Files\WPS Office\ksolaunch.exe",
    r"C:\Program Files (x86)\WPS Office\ksolaunch.exe",
    r"C:\Users\Public\WPS Office\ksolaunch.exe",
)


def _iter_word_exe_paths():
    """依次产出可能存在的 WINWORD.EXE 路径。"""
    guesses = list(_WINWORD_GUESSES)
    # 从 Program Files 目录动态发现 OfficeNN 目录
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if not base:
            continue
        office_root = os.path.join(base, "Microsoft Office", "Root")
        if os.path.isdir(office_root):
            try:
                for name in sorted(os.listdir(office_root), reverse=True):
                    if name.lower().startswith("office"):
                        guesses.append(os.path.join(office_root, name, "WINWORD.EXE"))
            except OSError:
                pass
        office_direct = os.path.join(base, "Microsoft Office")
        if os.path.isdir(office_direct):
            try:
                for name in sorted(os.listdir(office_direct), reverse=True):
                    if name.lower().startswith("office"):
                        guesses.append(os.path.join(office_direct, name, "WINWORD.EXE"))
            except OSError:
                pass
    seen = set()
    for guess in guesses:
        key = guess.lower()
        if key in seen:
            continue
        seen.add(key)
        yield guess


def find_word(verbose: list[str] | None = None) -> str | None:
    """查找本机 Microsoft Word 的 winword.exe 路径，找不到返回 None。

    verbose 传入列表时，会把每一步探测结果写进去，便于排查检测失败的原因。
    """
    def note(msg: str) -> None:
        if verbose is not None:
            verbose.append(msg)

    try:
        import winreg
    except ImportError:
        note("无法导入 winreg（当前不是 Windows 环境或非标准 Python）")
        return None

    candidates = (
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe",
         "HKLM App Paths"),
        (winreg.HKEY_CURRENT_USER,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe",
         "HKCU App Paths"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\winword.exe",
         "HKLM(32) App Paths"),
    )
    for root, sub, label in candidates:
        try:
            with winreg.OpenKey(root, sub) as key:
                value, _ = winreg.QueryValueEx(key, None)
                if value and os.path.isfile(value):
                    note(f"注册表 {label} 命中: {value}")
                    return value
                note(f"注册表 {label} 有值但文件不存在: {value!r}")
        except FileNotFoundError:
            note(f"注册表 {label} 不存在")
        except OSError as exc:
            note(f"注册表 {label} 读取失败: {exc}")

    for guess in _iter_word_exe_paths():
        if os.path.isfile(guess):
            note(f"按安装路径命中: {guess}")
            return guess
    note("注册表与常见安装路径均未找到 WINWORD.EXE")
    return None


def find_wps(verbose: list[str] | None = None) -> str | None:
    """查找 WPS Office（其文字组件兼容 Word COM 接口）。"""
    def note(msg: str) -> None:
        if verbose is not None:
            verbose.append(msg)

    for guess in _WPS_GUESSES:
        if os.path.isfile(guess):
            note(f"找到 WPS: {guess}")
            return guess
    for exe in ("wps.exe", "et.exe"):
        found = shutil.which(exe)
        if found:
            note(f"PATH 中找到 WPS: {found}")
            return found
    return None


def find_soffice(verbose: list[str] | None = None) -> str | None:
    """查找 LibreOffice 命令行程序。"""
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        if verbose is not None:
            verbose.append(f"PATH 中找到 LibreOffice: {found}")
        return found
    for guess in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if os.path.isfile(guess):
            if verbose is not None:
                verbose.append(f"找到 LibreOffice: {guess}")
            return guess
    if verbose is not None:
        verbose.append("未找到 LibreOffice")
    return None


def _pywin32_status() -> tuple[bool, str]:
    """返回 (是否可用, 失败原因)。"""
    try:
        import pythoncom  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return False, f"pythoncom 导入失败：{type(exc).__name__}: {exc}"
    try:
        import win32com.client  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return False, f"win32com.client 导入失败：{type(exc).__name__}: {exc}"
    return True, ""


def _has_pywin32() -> bool:
    return _pywin32_status()[0]


def _progid_registered(progid: str) -> bool:
    """检查 COM ProgID 是否已注册。"""
    try:
        import winreg
    except ImportError:
        return False
    for root in (winreg.HKEY_CLASSES_ROOT,):
        try:
            with winreg.OpenKey(root, progid):
                return True
        except OSError:
            continue
    return False


_engine_cache: dict[str, object] = {}


def clear_engine_cache() -> None:
    """清空引擎探测缓存（供界面「重新检测」使用）。"""
    _engine_cache.clear()


def word_engine_available() -> bool:
    """是否可用 Word/WPS COM 引擎。"""
    if "word" not in _engine_cache:
        _engine_cache["word"] = bool(find_word()) and _has_pywin32()
    return bool(_engine_cache["word"])


def com_engine_progid() -> str | None:
    """返回实际可用的 COM ProgID（Word 优先，其次 WPS）。"""
    if "progid" not in _engine_cache:
        _engine_cache["progid"] = None
        if word_engine_available():
            for progid in WORD_PROGIDS:
                if _progid_registered(progid):
                    _engine_cache["progid"] = progid
                    break
            else:
                _engine_cache["progid"] = WORD_PROGIDS[0]
        elif _has_pywin32():
            for progid in WPS_PROGIDS:
                if _progid_registered(progid):
                    _engine_cache["progid"] = progid
                    break
    return _engine_cache["progid"]  # type: ignore[return-value]


def describe_engine() -> str:
    """给界面显示当前 Word->PDF 引擎。"""
    if word_engine_available():
        exe = find_word() or ""
        version = ""
        m = re.search(r"Office(\d+)", exe)
        if m:
            version = f" Office {m.group(1)}"
        return f"Microsoft Word{version}（版式与 Word 完全一致）"
    progid = com_engine_progid()
    if progid:
        return f"WPS Office（{progid}）"
    soffice = find_soffice()
    if soffice:
        return "LibreOffice（未检测到 Word）"
    return "无可用引擎（需要 Microsoft Word 或 LibreOffice）"


def diagnose_engine() -> str:
    """生成一份完整的引擎检测报告，用于排查「未检测到 Word」类问题。"""
    lines: list[str] = []
    lines.append("===== 运行环境 =====")
    lines.append(f"Python      : {sys.version.split()[0]}")
    lines.append(f"解释器路径  : {sys.executable}")
    lines.append(f"程序目录    : {BASE_DIR}")
    lines.append(f"系统        : {platform.platform()}")

    lines.append("")
    lines.append("===== COM 支持（pywin32）=====")
    ok, why = _pywin32_status()
    lines.append("可用        : " + ("是" if ok else "否"))
    if not ok:
        lines.append("原因        : " + why)
        # 最常见的原因：用了系统 Python，而不是程序自带的 env
        env_dir = os.path.join(BASE_DIR, "env")
        env_py = os.path.join(env_dir, "Scripts", "python.exe")
        in_env = False
        try:
            in_env = os.path.normcase(os.path.abspath(sys.prefix)) == os.path.normcase(env_dir)
        except Exception:  # noqa: BLE001
            pass
        env_has_pywin32 = os.path.isdir(
            os.path.join(env_dir, "Lib", "site-packages", "pywin32_system32")
        )
        if not in_env and env_has_pywin32:
            lines.append("")
            lines.append("⚠ 当前用的不是程序自带的运行环境：")
            lines.append(f"    正在使用: {sys.executable}")
            lines.append(f"    应当使用: {env_py}")
            lines.append("  这正是 Word 引擎不可用的原因（pywin32 只装在自带的 env 里）。")
            lines.append("  解决办法：")
            lines.append("    1) 双击程序目录下的「启动程序.bat」（它会自动用自带环境）；")
            lines.append("    2) 或在终端里显式使用自带解释器，例如：")
            lines.append(f"       \"{env_py}\" cli.py --diagnose dummy")
        elif not env_has_pywin32:
            lines.append("")
            lines.append("⚠ 程序自带的 env 里没有 pywin32，运行环境可能不完整。")
            lines.append("  请双击「安装依赖.bat」重建运行环境。")

    lines.append("")
    lines.append("===== 查找 Microsoft Word =====")
    detail: list[str] = []
    found = find_word(detail)
    for item in detail:
        lines.append("  " + item)
    lines.append("结果        : " + (found if found else "未找到"))

    lines.append("")
    lines.append("===== 查找 WPS Office =====")
    wps_detail: list[str] = []
    wps = find_wps(wps_detail)
    for item in wps_detail:
        lines.append("  " + item)
    lines.append("结果        : " + (wps if wps else "未找到"))

    lines.append("")
    lines.append("===== 查找 LibreOffice =====")
    lo_detail: list[str] = []
    lo = find_soffice(lo_detail)
    for item in lo_detail:
        lines.append("  " + item)
    lines.append("结果        : " + (lo if lo else "未找到"))

    lines.append("")
    lines.append("===== COM ProgID 注册情况 =====")
    for progid in WORD_PROGIDS + WPS_PROGIDS:
        lines.append(f"  {progid:22s} : {'已注册' if _progid_registered(progid) else '未注册'}")

    lines.append("")
    lines.append("===== 结论 =====")
    lines.append("Word -> PDF 引擎: " + describe_engine())
    if not word_engine_available() and not com_engine_progid() and not lo:
        lines.append("")
        lines.append("没有可用引擎。解决办法（任选其一）：")
        lines.append("  1. 安装 Microsoft Word（推荐，效果最好）；")
        lines.append("  2. 安装 LibreOffice（免费，https://zh-cn.libreoffice.org/）；")
        lines.append("  3. 若你已装 WPS：打开一次 WPS 文字完成初始化，再点「重新检测」。")
    return "\n".join(lines)



# ---------------------------------------------------------------- Word -> PDF


@dataclass
class _WordSession:
    """一次 Word COM 会话；批量转换时复用，避免反复冷启动（冷启动约 10 秒）。"""

    app: object = None
    progid: str = ""

    def start(self) -> None:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        progid = com_engine_progid() or WORD_PROGIDS[0]
        self.progid = progid
        self.app = win32com.client.DispatchEx(progid)
        for attr, value in (("Visible", False), ("DisplayAlerts", 0)):
            try:
                setattr(self.app, attr, value)
            except Exception:  # noqa: BLE001
                pass
        for name, value in (
            ("WarnBeforeSavingPrintingSendingMarkup", False),
            ("ConfirmConversions", False),
        ):
            try:
                setattr(self.app.Options, name, value)
            except Exception:  # noqa: BLE001
                pass
        try:
            self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:  # noqa: BLE001
            pass

    def convert(self, src: str, dst: str, progress: ProgressCb | None = None) -> None:
        if self.app is None:
            if progress:
                progress("正在启动转换引擎…")
            self.start()
        doc = None
        try:
            if progress:
                progress("正在打开文档…")
            doc = self.app.Documents.Open(
                os.path.abspath(src),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            if progress:
                progress("正在排版并导出 PDF…")
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(dst),
                ExportFormat=WD_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=0,        # wdExportOptimizeForPrint
                Range=0,              # wdExportAllDocument
                Item=0,               # wdExportDocumentContent
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,    # wdExportCreateHeadingBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
            )
            if progress:
                progress("导出完成")
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:  # noqa: BLE001
                    pass

    def stop(self) -> None:
        if self.app is not None:
            try:
                self.app.Quit(SaveChanges=0)
            except Exception:  # noqa: BLE001
                pass
            self.app = None
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


_EVENT_ERROR_HINT = (
    "Word 未能响应自动化调用。常见原因：\n"
    "  • Word 正开着对话框或正被打断（请手动打开再关闭一次 Word）；\n"
    "  • Word 处于「首次运行/激活」状态，需要先手动完成一次启动；\n"
    "  • 当前进程权限受限（例如在受限沙箱或服务中运行）。"
)


def _is_com_unavailable(msg: str) -> bool:
    """判断是否属于「COM 组件调用不起来」这类错误。"""
    lowered = msg.lower()
    return (
        "未能引发事件" in msg
        or "-2146822286" in msg
        or "could not raise" in lowered
        or "-2147221005" in msg          # 类字符串无效：ProgID 未注册
        or "invalid class string" in lowered
        or "class not registered" in lowered
        or "库未注册" in msg
        or "0x80040154" in lowered
    )


def _word2pdf_com(src: str, dst: str, session: _WordSession | None = None,
                  progress: ProgressCb | None = None,
                  allow_fallback: bool = True) -> None:
    """用 COM 引擎导出 PDF；COM 不可用且装了 LibreOffice 时自动回退。"""
    def fallback_or_raise(exc: Exception) -> None:
        msg = str(exc)
        if allow_fallback and find_soffice():
            if progress:
                progress("COM 引擎不可用，改用 LibreOffice 重试…")
            _word2pdf_soffice(src, dst, progress)
            return
        if _is_com_unavailable(msg):
            raise ConvertError(
                _EVENT_ERROR_HINT
                + "\n\n本次使用的 COM 组件为："
                + (com_engine_progid() or "未知")
                + "\n可点界面上方的「诊断报告」查看详细探测结果。"
            ) from exc
        raise ConvertError(f"Word 转换失败：{msg}") from exc

    if session is not None:
        try:
            session.convert(src, dst, progress)
        except Exception as exc:  # noqa: BLE001
            fallback_or_raise(exc)
        return
    own = _WordSession()
    try:
        own.convert(src, dst, progress)   # start() 内部会按需启动
    except ConvertError:
        raise
    except Exception as exc:  # noqa: BLE001
        fallback_or_raise(exc)
    finally:
        own.stop()


def _word2pdf_soffice(src: str, dst: str, progress: ProgressCb | None = None) -> None:
    soffice = find_soffice()
    if not soffice:
        detail: list[str] = []
        find_word(detail)
        raise ConvertError(
            "本机未检测到可用的转换引擎。\n"
            + "\n".join("  · " + item for item in detail)
            + "\n请安装以下任一项后重试：\n"
            "  · Microsoft Word（推荐，版式最准确）\n"
            "  · LibreOffice（免费，https://zh-cn.libreoffice.org/）\n"
            "若你已装 WPS Office，请先打开一次 WPS 文字完成初始化。\n"
            "详细检测报告：点击界面上方的「诊断报告」按钮。"
        )
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    os.makedirs(outdir, exist_ok=True)
    cmd = [
        soffice, "--headless", "--norestore", "--invisible",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", outdir, os.path.abspath(src),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if not os.path.isfile(produced):
        raise ConvertError(
            "LibreOffice 转换失败：" + (proc.stderr or proc.stdout or "无输出").strip()[:400]
        )
    if os.path.abspath(produced) != os.path.abspath(dst):
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(produced, dst)


def word_to_pdf(src: str, dst: str, *, session: "_WordSession | None" = None,
                progress: ProgressCb | None = None,
                engine: str = "auto") -> str:
    """把 Word 文档转换为 PDF，返回输出路径。

    engine 可选：
        "auto"        —— 有 Word/WPS 就用它，否则用 LibreOffice（默认）
        "office"      —— 强制走 COM（Word / WPS）
        "libreoffice" —— 强制走 LibreOffice（用于部署到无 Word 的服务器）
    """
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise ConvertError(f"源文件不存在：{src}")
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)

    # 引擎判定只做一次，避免下面分支各自重复探测导致结论不一致
    progid = com_engine_progid()
    com_ready = word_engine_available() or bool(progid)
    soffice = find_soffice()

    if engine == "office":
        use_com = True
    elif engine == "libreoffice":
        use_com = False
    else:
        use_com = com_ready

    if use_com:
        if not com_ready:
            # 被强制要求走 Office 却没有引擎，直接给出明确原因
            _word2pdf_soffice(src, dst, progress)
            return dst
        label = "WPS Office" if (progid and progid != "Word.Application") else "Word"
        if progress:
            progress(f"正在调用 {label} 引擎导出：{os.path.basename(src)}")
        _word2pdf_com(src, dst, session, progress)
    elif soffice:
        if progress:
            progress(f"正在调用 LibreOffice 导出：{os.path.basename(src)}")
        _word2pdf_soffice(src, dst, progress)
    else:
        # 没有任何引擎：给出可诊断的明确提示（不再报模糊错误）
        _word2pdf_soffice(src, dst, progress)

    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        raise ConvertError("引擎未生成 PDF 文件。")
    return dst


# ---------------------------------------------------------------- PDF -> Word

_CJK_RE = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uff00-\uffef]")
_HEADING_RE = re.compile(
    r"^\s*(?:第\s*[0-9一二三四五六七八九十百]+\s*[章节篇部分]"
    r"|[0-9]+(?:\.[0-9]+){1,3}\s*[、.．]?\s*\S"
    r"|[一二三四五六七八九十]+[、.．]\s*\S"
    r"|Chapter\s+\d+|ABSTRACT|摘\s*要|前\s*言|引\s*言|结\s*论|参考文献)\s*",
    re.IGNORECASE,
)


@dataclass
class PdfToWordOptions:
    """PDF 转 Word 的可调参数。"""

    detect_headings: bool = True   # 依据字号/加粗/编号推断标题层级
    keep_images: bool = True       # 提取并嵌入页面图片
    keep_tables: bool = True       # 识别表格并还原为 Word 表格
    page_break: bool = True        # 每页之间插入分页符
    scanned_fallback: bool = True  # 无文字层（扫描件）时整页转图片
    image_dpi: int = 150           # 扫描件整页渲染精度
    progress: ProgressCb | None = None
    cancel: CancelCb | None = None


@dataclass
class PdfToWordResult:
    output: str
    pages: int = 0
    paragraphs: int = 0
    tables: int = 0
    images: int = 0
    scanned: bool = False
    warnings: list[str] = field(default_factory=list)


def _load_pymupdf():
    try:
        import pymupdf

        return pymupdf
    except ImportError:
        import fitz as pymupdf  # type: ignore

        return pymupdf


def _span_flags(span: dict) -> tuple[bool, bool]:
    font = (span.get("font") or "").lower()
    flags = int(span.get("flags", 0) or 0)
    bold = ("bold" in font) or ("black" in font) or ("heavy" in font) or bool(flags & 16)
    italic = ("italic" in font) or ("oblique" in font) or bool(flags & 2)
    return bold, italic


def _line_text_and_style(line: dict) -> tuple[str, float, float]:
    """合成一行的文本，返回 (文本, 加权平均字号, 粗体字符占比)。"""
    parts: list[str] = []
    sizes: list[float] = []
    bold_chars = 0
    total = 0
    for span in line.get("spans", []):
        text = span.get("text", "")
        if not text:
            continue
        parts.append(text)
        stripped = text.strip()
        if stripped:
            n = len(stripped)
            size = float(span.get("size", 0) or 0)
            sizes.extend([size] * n)
            total += n
            bold, _ = _span_flags(span)
            if bold:
                bold_chars += n
    avg = (sum(sizes) / len(sizes)) if sizes else 0.0
    bold_ratio = (bold_chars / total) if total else 0.0
    return "".join(parts), avg, bold_ratio


_BULLET_PREFIXES = ("•", "·", "▪", "▫", "◦", "‣", "●", "○", "◆", "◇", "■", "□", "–", "—")


def _clean_line(text: str) -> str:
    """去掉行首项目符号字符，返回 (纯文本, 是否列表项)。

    PDF 里的「• 内容」常被拆成「•」和「内容」两行，这里统一识别。
    """
    stripped = text.strip()
    if not stripped:
        return "", False
    for bullet in _BULLET_PREFIXES:
        if stripped.startswith(bullet):
            return stripped[len(bullet):].strip(), True
    # 「- xxx」「* xxx」也算列表项，但避免把正文里的破折号误判
    if len(stripped) > 1 and stripped[0] in "-*" and stripped[1] == " ":
        return stripped[2:].strip(), True
    return stripped, False


def _join_lines(texts: list[str]) -> str:
    """拼接同一个段落的多行。

    中文之间直接相连；西文单词之间在换行处补一个空格，避免出现 "thatcontinues"。
    """
    out = ""
    for text in texts:
        if not text:
            continue
        if not out:
            out = text
            continue
        prev_ch = out[-1]
        next_ch = text[0]
        need_space = (
            (prev_ch.isascii() and prev_ch.isalnum())
            and (next_ch.isascii() and next_ch.isalnum())
        )
        out = out + (" " if need_space else "") + text
    return out.strip()


def _make_paragraph(lines: list[dict]) -> dict:
    """把若干行组合成一个「段落」对象。"""
    text = _join_lines([l["text"] for l in lines])
    sizes = [l["size"] for l in lines if l["size"] > 0]
    avg_size = (sum(sizes) / len(sizes)) if sizes else 0.0
    bold_lines = sum(1 for l in lines if l.get("bold"))
    x0 = min(l["x0"] for l in lines)
    x1 = max(l["x1"] for l in lines)
    y0 = min(l["y0"] for l in lines)
    y1 = max(l["y1"] for l in lines)
    return {
        "text": text,
        "avg_size": avg_size,
        "bold": bold_lines >= max(1, len(lines) // 2),
        "bbox": (x0, y0, x1, y1),
        "y": y0,
        "x": x0,
        # 构成该段落的每一行，用于判断后续行是否内缩（续行）
        "lines": [{"text": l["text"], "x0": l["x0"], "x1": l["x1"], "y1": l["y1"],
                   "size": l["size"], "is_bullet": l.get("is_bullet", False)}
                  for l in lines],
        "is_bullet": lines[0].get("is_bullet", False),
    }


def _page_body_size(blocks: list[dict]) -> float:
    """页面上出现次数最多的字号，视为正文字号。"""
    sizes = [round(b["avg_size"], 1) for b in blocks if b["avg_size"] > 0]
    if not sizes:
        return 0.0
    from collections import Counter

    return Counter(sizes).most_common(1)[0][0]


def _page_left_margin(blocks: list[dict], body_size: float) -> float:
    """页面版心左边距：正文字号行的最小左边界。"""
    xs = [l["x0"] for b in blocks
          if body_size <= 0 or b["avg_size"] <= body_size * 1.1
          for l in b.get("lines", [])]
    if not xs:
        xs = [b["bbox"][0] for b in blocks]
    return min(xs) if xs else 0.0


def _can_append(prev: dict, line: dict, body_margin: float, is_heading) -> bool:
    """判断行 line 是否应并入前一段落 prev。"""
    last = prev["lines"][-1]
    size = line["size"] or prev["avg_size"]
    line_h = max(size, 1.0)

    # 只含项目符号的行：它唯一的合法归宿是紧随其后的那一行内容，
    # 绝不允许把后面的文字并进一个空符号行。
    if last.get("is_bullet") and not last["text"].strip():
        if line.get("is_bullet"):
            return False
        return abs(line["y0"] - last["y1"]) < line_h * 2.5
    # 段内再次出现项目符号 -> 这是新的列表项，另起一段
    if line.get("is_bullet"):
        return False

    # 标题自成一段：绝不允许把后续内容并进标题
    if prev.get("is_heading"):
        return False

    # 标题（含与正文同字号的编号式标题）自成一段
    if is_heading(line["text"], size, prev["avg_size"]):
        return False

    gap = line["y0"] - last["y1"]
    if gap < 0:                      # 同一基线上的并列文本（表格/分栏残留）
        return False
    if gap > line_h * 1.15:          # 段间距明显变大 -> 新段落
        return False
    # 字号差异过大 -> 新段落（标题、图注等）
    if prev["avg_size"] > 0 and abs(line["size"] - prev["avg_size"]) > 1.6:
        return False

    cur_x = line["x0"]
    first_x = prev["lines"][0]["x0"]

    # 关键：首行缩进是相对「版心左边距」判断的。
    #   缩进行 + 回到版心的行  -> 同段换行（判断依据：当前行更靠左）
    #   缩进行 + 另一缩进行    -> 两个独立段落
    if cur_x < first_x - 4.0:
        return True
    if cur_x > first_x + 8.0:
        # 进一步内缩：若前面已经出现过回到版心的行，说明这是新段落
        if any(l["x0"] <= body_margin + 2.0 for l in prev["lines"]):
            return False
        return True
    if abs(cur_x - first_x) <= 6.0:
        # 左边界一致：只有当本段首行是「行首缩进」时才可能是同段换行
        cjk = bool(_CJK_RE.search(prev["text"]))
        expected = line_h * (2.0 if cjk else 1.5)
        return (first_x - body_margin) >= expected * 0.5
    return False


def _merge_and_normalize(blocks: list[dict], body_size: float = 0.0,
                         body_margin: float = 0.0,
                         is_heading=None) -> list[dict]:
    """把 PDF 的行块合并成真正的段落。

    PDF 没有「段落」概念：一段被自动换行拆成多行的文字会变成多个独立块。
    这里按「相对版心左边距的缩进 + 行距 + 字号」判断续行，合并回一段。
    """
    if is_heading is None:
        def is_heading(text: str, size: float, prev_size: float) -> bool:  # noqa: ARG001
            return False

    def _is_heading_line(text: str, size: float, prev_size: float) -> bool:
        # 与上一行字号接近时不做标题判断，避免正文被误判
        if prev_size > 0 and abs(size - prev_size) <= 1.6 and _HEADING_RE.match(text or ""):
            return True
        return bool(is_heading(text, size, prev_size))

    para_lines: list[list[dict]] = []
    for block in blocks:
        for line in block.get("lines", []):
            text = line["text"]
            if not text.strip():
                continue
            cleaned, is_bullet = _clean_line(text)
            entry = dict(line, text=cleaned, is_bullet=is_bullet)
            if not cleaned and not is_bullet:
                continue
            # 纯符号行（内容在下一行）没有需要吸附的上一段时，直接丢掉
            if not cleaned and is_bullet and not para_lines:
                continue
            if para_lines:
                prev = _make_paragraph(para_lines[-1])
                if _can_append(prev, entry, body_margin, _is_heading_line):
                    para_lines[-1].append(entry)
                    continue
            para_lines.append([entry])
    result = []
    for lines in para_lines:
        para = _make_paragraph(lines)
        if not para["text"]:
            continue
        if _is_heading_line(para["text"], para["avg_size"], para["avg_size"]):
            para["is_heading"] = True
        result.append(para)
    return result


def _split_columns(blocks: list[dict], page_width: float) -> list[list[dict]]:
    """按 x 方向投影判断分栏；单栏时返回一个分组。"""
    if len(blocks) < 6:
        return [blocks]
    buckets = [0] * 40
    for block in blocks:
        x0, _, x1, _ = block["bbox"]
        x0 = max(0.0, min(x0, page_width - 1))
        x1 = max(0.0, min(x1, page_width - 1))
        start = int(x0 / page_width * 40)
        end = int(x1 / page_width * 40)
        for b in range(start, end + 1):
            buckets[b] += 1
    threshold = max(1, len(blocks) // 6)
    empty = [i for i, count in enumerate(buckets) if count <= threshold]
    # 中部存在明显空白带 -> 判为分栏
    mid = [i for i in empty if 12 <= i <= 27]
    if not mid:
        return [blocks]
    split = (mid[0] + mid[-1]) / 2 / 40 * page_width
    left = [b for b in blocks if b["bbox"][0] < split]
    right = [b for b in blocks if b["bbox"][0] >= split]
    if not left or not right:
        return [blocks]
    return [left, right]


def _order_lines(lines: list[dict], tol: float = 3.5) -> list[dict]:
    """按视觉行分组后再按 x 排序，返回阅读顺序的行列表。

    同一行里各 span 的 y 坐标可能相差数点（项目符号与正文尤其明显），
    因此先用容差把行聚成「视觉行」，行内再按 x 排序。
    """
    if not lines:
        return []
    ordered: list[dict] = []
    rows: list[list[dict]] = []
    for line in sorted(lines, key=lambda l: l["y0"]):
        if rows and line["y0"] - rows[-1][0]["y0"] <= tol:
            rows[-1].append(line)
        else:
            rows.append([line])
    for row in rows:
        row.sort(key=lambda l: l["x0"])
        ordered.extend(row)
    return ordered


def _page_blocks(page) -> list[dict]:
    """把页面文本整理成按阅读顺序排列、并已合并换行的段落块。"""
    blocks: list[dict] = []
    raw = page.get_text("dict", sort=True)
    for block in raw.get("blocks", []):
        if block.get("type") != 0:
            continue
        lines = []
        for line in block.get("lines", []):
            text, size, bold_ratio = _line_text_and_style(line)
            if not text.strip():
                continue
            lbbox = line.get("bbox", (0, 0, 0, 0))
            lines.append({
                "text": text,
                "size": size,
                "bold": bold_ratio > 0.6,
                "x0": lbbox[0],
                "x1": lbbox[2],
                "y0": lbbox[1],
                "y1": lbbox[3],
            })
        if not lines:
            continue
        # 同一视觉行内的项目符号与文字，其 y 坐标会有细微差异（如 343.55 / 342.97），
        # 若直接「先按 y 排序」会把符号行排到文字行之后。这里先按基线聚类成行，
        # 再按 x 排序，保证「• 内容」的顺序正确。
        lines = _order_lines(lines)
        sizes = [l["size"] for l in lines if l["size"] > 0]
        avg_size = (sum(sizes) / len(sizes)) if sizes else 0.0
        bold_lines = sum(1 for l in lines if l["bold"])
        bbox = tuple(block.get("bbox", (0, 0, 0, 0)))
        blocks.append({
            "type": "text",
            "text": " ".join(l["text"].strip() for l in lines if l["text"].strip()),
            "avg_size": avg_size,
            "bold": bold_lines >= max(1, len(lines) // 2),
            "bbox": bbox,
            "y": bbox[1],
            "x": bbox[0],
            "lines": lines,
        })

    blocks.sort(key=lambda b: (round(b["y"], 1), round(b["x"], 1)))
    page_body_size = _page_body_size(blocks)
    page_margin = _page_left_margin(blocks, page_body_size)

    def is_heading(text: str, size: float, prev_size: float) -> bool:
        # 与上一行同字号时，靠「编号 + 长度」判断（需要加粗或编号特征）
        if prev_size > 0 and abs(size - prev_size) <= 1.6:
            return _heading_level(text, size, page_body_size, False, True) > 0
        return size >= page_body_size * 1.35

    ordered: list[dict] = []
    for column in _split_columns(blocks, page.rect.width):
        column.sort(key=lambda b: (round(b["y"], 1), round(b["x"], 1)))
        col_body = _page_body_size(column)
        col_margin = _page_left_margin(column, col_body)
        ordered.extend(_merge_and_normalize(
            column,
            body_size=col_body,
            body_margin=col_margin,
            is_heading=is_heading,
        ))
    return ordered


def _heading_level(text: str, avg_size: float, body_size: float, bold: bool,
                   detect: bool) -> int:
    """返回 0 表示正文，1-4 表示标题层级。"""
    if not detect:
        return 0
    stripped = text.strip()
    if not stripped or len(stripped) > 80:
        return 0
    if body_size > 0:
        if avg_size >= body_size * 1.8:
            return 1
        if avg_size >= body_size * 1.55:
            return 2
        if avg_size >= body_size * 1.35:
            return 3
        if avg_size >= body_size * 1.12 and bold:
            return 4
    if _HEADING_RE.match(stripped) and len(stripped) <= 60:
        if re.match(r"^\s*第\s*[0-9一二三四五六七八九十百]+\s*[章篇]", stripped):
            return 1
        if re.match(r"^\s*(摘\s*要|前\s*言|引\s*言|结\s*论|参考文献|ABSTRACT)", stripped, re.I):
            return 1
        depth = stripped.count(".")
        return min(2 + depth, 4)
    if bold and len(stripped) <= 24 and not stripped.endswith(("。", "，", ".", ",")):
        return 4
    return 0


def _write_paragraph_text(paragraph, text: str) -> None:
    """把可能含换行的文本写入段落，用软回车保留原换行。"""
    for i, chunk in enumerate(text.split("\n")):
        if i:
            paragraph.add_run().add_break()
        if chunk:
            paragraph.add_run(chunk)


def pdf_to_word(src: str, dst: str, options: PdfToWordOptions | None = None) -> PdfToWordResult:
    """把 PDF 转换为可编辑的 Word 文档。"""
    opts = options or PdfToWordOptions()
    pymupdf = _load_pymupdf()
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt

    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.isfile(src):
        raise ConvertError(f"源文件不存在：{src}")
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)

    result = PdfToWordResult(output=dst)
    doc = docx.Document()

    # 中文字体与字号
    try:
        normal = doc.styles["Normal"]
        normal.font.name = "宋体"
        normal.font.size = Pt(10.5)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    except Exception:  # noqa: BLE001
        pass

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(25)
    section.right_margin = Mm(25)
    section.top_margin = Mm(25)
    section.bottom_margin = Mm(25)

    workdir = tempfile.mkdtemp(prefix="pdf2word_", dir=temp_root())
    try:
        with pymupdf.open(src) as pdf:
            if pdf.needs_pass:
                raise ConvertError("该 PDF 有密码保护，请先解除密码再转换。")
            result.pages = pdf.page_count
            if result.pages == 0:
                raise ConvertError("该 PDF 不包含任何页面。")

            for pno in range(pdf.page_count):
                if opts.cancel and opts.cancel():
                    raise Cancelled("已取消")
                page = pdf.load_page(pno)
                if opts.progress:
                    opts.progress(f"解析第 {pno + 1}/{pdf.page_count} 页")

                raw_text = page.get_text("text").strip()
                # 扫描件（没有文字层）整页转成图片，保证内容不丢失
                if len(raw_text) < 15 and opts.scanned_fallback:
                    result.scanned = True
                    pix = page.get_pixmap(dpi=opts.image_dpi)
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    para.add_run().add_picture(io.BytesIO(pix.tobytes("png")), width=Mm(160))
                    result.images += 1
                    if pno < pdf.page_count - 1 and opts.page_break:
                        doc.add_page_break()
                    continue

                table_rects: list = []
                if opts.keep_tables:
                    try:
                        for tbl in (page.find_tables().tables or []):
                            data = tbl.extract()
                            if not data or not any(
                                any((c or "").strip() for c in row) for row in data
                            ):
                                continue
                            table_rects.append(pymupdf.Rect(tbl.bbox))
                            rows = len(data)
                            cols = max(len(r) for r in data)
                            word_tbl = doc.add_table(rows=rows, cols=cols)
                            word_tbl.style = "Table Grid"
                            for ri, row in enumerate(data):
                                for ci in range(cols):
                                    cell = (row[ci] if ci < len(row) else "") or ""
                                    word_tbl.cell(ri, ci).text = cell.replace("\n", " ").strip()
                            result.tables += 1
                    except Exception as exc:  # noqa: BLE001
                        result.warnings.append(f"第 {pno + 1} 页表格识别失败：{exc}")

                def inside_table(bbox) -> bool:
                    if not bbox:
                        return False
                    r = pymupdf.Rect(bbox)
                    area = max(r.get_area(), 1e-6)
                    return any((r & tr).get_area() > 0.6 * area for tr in table_rects)

                blocks = _page_blocks(page)
                sizes = [b["avg_size"] for b in blocks
                         if b["avg_size"] > 0 and not inside_table(b["bbox"])]
                body_size = 0.0
                if sizes:
                    from collections import Counter

                    body_size = Counter(round(s, 1) for s in sizes).most_common(1)[0][0]

                for block in blocks:
                    if inside_table(block["bbox"]):
                        continue
                    text = block["text"].strip()
                    if not text:
                        continue
                    level = _heading_level(text, block["avg_size"], body_size,
                                           block["bold"], opts.detect_headings)
                    if level:
                        heading = doc.add_heading(level=min(level, 4))
                        _write_paragraph_text(heading, text)
                        result.paragraphs += 1
                        continue
                    try:
                        para = doc.add_paragraph(style="List Bullet" if block.get("is_bullet")
                                                 else None)
                    except Exception:  # noqa: BLE001
                        para = doc.add_paragraph()
                    _write_paragraph_text(para, text)
                    bbox = block["bbox"]
                    if (bbox and bbox[0] > page.rect.width * 0.4 and len(text) < 40
                            and not block.get("is_bullet")):
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    result.paragraphs += 1

                if opts.keep_images:
                    for img_index, info in enumerate(page.get_images(full=True)):
                        if opts.cancel and opts.cancel():
                            raise Cancelled("已取消")
                        try:
                            base = pdf.extract_image(info[0])
                        except Exception:  # noqa: BLE001
                            continue
                        if not base or not base.get("image"):
                            continue
                        if base.get("width", 0) < 24 or base.get("height", 0) < 24:
                            continue
                        try:
                            para = doc.add_paragraph()
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            para.add_run().add_picture(io.BytesIO(base["image"]), width=Mm(150))
                            result.images += 1
                        except Exception as exc:  # noqa: BLE001
                            result.warnings.append(f"第 {pno + 1} 页图片嵌入失败：{exc}")

                if pno < pdf.page_count - 1 and opts.page_break:
                    doc.add_page_break()

        if opts.progress:
            opts.progress("正在写出 Word 文件")
        doc.save(dst)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        raise ConvertError("未能生成 Word 文件。")
    return result


# ---------------------------------------------------------------- 统一入口


@dataclass
class Job:
    """一个待转换任务。"""

    src: str
    dst: str
    direction: str  # 'word2pdf' | 'pdf2word'


def default_output(src: str, outdir: str, direction: str) -> str:
    """给出不覆盖已有文件的默认输出路径。"""
    stem = os.path.splitext(os.path.basename(src))[0]
    suffix = ".pdf" if direction == "word2pdf" else ".docx"
    dst = os.path.join(outdir, stem + suffix)
    if os.path.abspath(dst).lower() == os.path.abspath(src).lower():
        dst = os.path.join(outdir, stem + "_converted" + suffix)
    base, ext = os.path.splitext(dst)
    n = 1
    while os.path.exists(dst):
        dst = f"{base}({n}){ext}"
        n += 1
    return dst


def run_jobs(jobs: Sequence[Job], *, options: PdfToWordOptions | None = None,
             progress: ProgressCb | None = None,
             cancel: CancelCb | None = None,
             engine: str = "auto") -> list[tuple[Job, str | None]]:
    """批量执行任务，返回 [(任务, 错误信息或 None)]。

    Word 会话在整批 Word->PDF 任务间复用，批量转换更快。

    engine: "auto" / "office" / "libreoffice"
        —— "libreoffice" 用于部署到没有 Word 的服务器（例如手机版后端）。
    """
    outcomes: list[tuple[Job, str | None]] = []
    session = _WordSession()
    session_started = False
    use_com = engine == "office" or (engine == "auto" and word_engine_available())
    try:
        for index, job in enumerate(jobs, 1):
            if cancel and cancel():
                outcomes.append((job, "已取消"))
                continue
            name = os.path.basename(job.src)
            if progress:
                progress(f"[{index}/{len(jobs)}] {name}")
            try:
                if job.direction == "word2pdf":
                    if use_com:
                        if not session_started:
                            if progress:
                                progress("正在启动 Word 引擎（首次启动约 10 秒）")
                            session.start()
                            session_started = True
                        word_to_pdf(job.src, job.dst, session=session,
                                    progress=progress, engine=engine)
                    else:
                        word_to_pdf(job.src, job.dst, progress=progress, engine=engine)
                else:
                    opts = options or PdfToWordOptions()
                    opts.progress = progress
                    opts.cancel = cancel
                    pdf_to_word(job.src, job.dst, opts)
                outcomes.append((job, None))
            except Cancelled:
                outcomes.append((job, "已取消"))
            except ConvertError as exc:
                outcomes.append((job, str(exc)))
            except Exception as exc:  # noqa: BLE001
                outcomes.append((job, f"{type(exc).__name__}: {exc}"))
    finally:
        if session_started:
            session.stop()
    return outcomes
