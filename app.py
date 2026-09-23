"""Word <-> PDF 互转工具 — 图形界面。

运行：双击「启动程序.bat」，或执行 env\\Scripts\\pythonw.exe app.py
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, X, Y, filedialog, messagebox
import tkinter as tk
from tkinter import ttk

APP_TITLE = "Word ⇄ PDF 互转工具"
APP_VERSION = "1.0"

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


def _bootstrap_env() -> None:
    """确保使用程序自带的 env 运行环境。

    如果被系统 Python 直接启动（例如在终端里敲 python app.py），
    会自动换成程序自带的 env 重新启动自己。否则会因为系统 Python 缺少
    pywin32 而误报「未检测到 Microsoft Word」。
    """
    if os.environ.get("WPC_REEXEC") == "1":
        return
    env_dir = os.path.join(HERE, "env")
    # 已经跑在自带的 env 里？直接返回
    try:
        if os.path.normcase(os.path.abspath(sys.prefix)) == os.path.normcase(env_dir):
            return
    except Exception:  # noqa: BLE001
        pass
    # env 的 site-packages 已在搜索路径里，也没问题
    if any("site-packages" in p and os.path.normcase(env_dir) in os.path.normcase(p)
           for p in sys.path):
        return

    name = "pythonw.exe" if os.path.basename(sys.executable).lower().startswith("pythonw") \
        else "python.exe"
    venv_py = os.path.join(env_dir, "Scripts", name)
    if not os.path.isfile(venv_py):
        venv_py = os.path.join(env_dir, "Scripts", "python.exe")
    if not os.path.isfile(venv_py):
        return   # 没有自带环境，只能用当前解释器

    # 自带环境里必须真的装了 pywin32，否则切过去反而更糟
    site = os.path.join(env_dir, "Lib", "site-packages")
    system32 = os.path.join(site, "pywin32_system32")
    if not os.path.isdir(system32):
        return
    if not any(n.lower().startswith("pywintypes") and n.lower().endswith(".dll")
               for n in os.listdir(system32)):
        return

    os.environ["WPC_REEXEC"] = "1"
    try:
        os.execv(venv_py, [venv_py, os.path.abspath(__file__), *sys.argv[1:]])
    except Exception:  # noqa: BLE001
        os.environ.pop("WPC_REEXEC", None)


_bootstrap_env()

import converter  # noqa: E402

# 可选：拖拽支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_OK = True
except Exception:  # noqa: BLE001
    DND_FILES = None
    TkinterDnD = None
    _DND_OK = False


# ---------------------------------------------------------------- 主窗口


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("880x620")
        self.root.minsize(760, 540)

        self.files: list[str] = []
        self.outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.same_as_source = tk.BooleanVar(value=False)
        self.direction = tk.StringVar(value="auto")

        self.opt_headings = tk.BooleanVar(value=True)
        self.opt_images = tk.BooleanVar(value=True)
        self.opt_tables = tk.BooleanVar(value=True)
        self.opt_pagebreak = tk.BooleanVar(value=True)
        self.opt_scanned = tk.BooleanVar(value=True)

        self.msg_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_flag = threading.Event()

        self._build_ui()
        self._refresh_engine_label()
        converter.cleanup_temp()          # 清理上次异常退出可能残留的临时文件
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(80, self._pump)
        self._bind_dnd()

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("正在转换", "还有转换任务正在进行，确定要退出吗？"):
                return
            self.cancel_flag.set()
        converter.cleanup_temp()
        self.root.destroy()

    # ------------------------------------------------------------ 界面

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill=BOTH, expand=True)

        # 顶部：引擎状态
        top = ttk.LabelFrame(outer, text="转换引擎", padding=(10, 6))
        top.pack(fill=X)
        self.engine_label = ttk.Label(top, text="检测中…", foreground="#0a5")
        self.engine_label.pack(side=LEFT)
        ttk.Button(top, text="重新检测", width=10, command=self._refresh_engine_label).pack(side=RIGHT)
        ttk.Button(top, text="诊断报告", width=10,
                   command=self._show_diagnostics).pack(side=RIGHT, padx=(0, 6))
        ttk.Button(top, text="转换说明", width=10, command=self._show_help).pack(side=RIGHT, padx=(0, 6))

        # 中部左：文件列表
        mid = ttk.Frame(outer)
        mid.pack(fill=BOTH, expand=True, pady=(10, 0))

        left = ttk.LabelFrame(mid, text="待转换文件（可拖拽文件到此处）", padding=(8, 6))
        left.pack(side=LEFT, fill=BOTH, expand=True)

        list_wrap = ttk.Frame(left)
        list_wrap.pack(fill=BOTH, expand=True)
        self.listbox = tk.Listbox(list_wrap, selectmode=tk.EXTENDED, activestyle="none",
                                  font=("Microsoft YaHei UI", 10))
        scroll = ttk.Scrollbar(list_wrap, orient=VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)

        btns = ttk.Frame(left)
        btns.pack(fill=X, pady=(8, 0))
        ttk.Button(btns, text="添加文件", command=self.add_files).pack(side=LEFT)
        ttk.Button(btns, text="添加文件夹", command=self.add_folder).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(side=LEFT)
        ttk.Button(btns, text="清空", command=self.clear_files).pack(side=LEFT, padx=6)

        # 中部右：转换方向
        right = ttk.LabelFrame(mid, text="转换方向", padding=(8, 6))
        right.pack(side=RIGHT, fill=Y, padx=(10, 0))
        for value, text in (
            ("auto", "自动识别（推荐）"),
            ("word2pdf", "Word → PDF"),
            ("pdf2word", "PDF → Word"),
        ):
            ttk.Radiobutton(right, text=text, value=value,
                            variable=self.direction, command=self._update_counts).pack(anchor="w", pady=2)

        ttk.Separator(right, orient="horizontal").pack(fill=X, pady=8)
        ttk.Label(right, text="PDF → Word 选项", font=("Microsoft YaHei UI", 9, "bold")).pack(anchor="w")
        for var, text in (
            (self.opt_headings, "识别标题层级"),
            (self.opt_images, "提取图片"),
            (self.opt_tables, "还原表格"),
            (self.opt_pagebreak, "按页分页"),
            (self.opt_scanned, "扫描件转图片"),
        ):
            ttk.Checkbutton(right, text=text, variable=var).pack(anchor="w")

        # 输出目录
        out = ttk.LabelFrame(outer, text="输出位置", padding=(10, 6))
        out.pack(fill=X, pady=(10, 0))
        self.out_entry = ttk.Entry(out, textvariable=self.outdir)
        self.out_entry.pack(side=LEFT, fill=X, expand=True)
        ttk.Button(out, text="浏览…", width=8, command=self.choose_outdir).pack(side=LEFT, padx=(6, 0))
        ttk.Checkbutton(out, text="与源文件同目录",
                        variable=self.same_as_source,
                        command=self._toggle_outdir).pack(side=LEFT, padx=(8, 0))

        # 进度与日志
        prog = ttk.LabelFrame(outer, text="进度", padding=(10, 6))
        prog.pack(fill=X, pady=(10, 0))
        self.progress = ttk.Progressbar(prog, mode="determinate")
        self.progress.pack(fill=X)
        self.status = ttk.Label(prog, text="就绪", foreground="#444")
        self.status.pack(anchor="w", pady=(4, 0))

        logf = ttk.LabelFrame(outer, text="日志", padding=(6, 4))
        logf.pack(fill=BOTH, expand=True, pady=(10, 0))
        self.log = tk.Text(logf, height=7, wrap="word", font=("Consolas", 9),
                           background="#1e1e1e", foreground="#dcdcdc", insertbackground="#dcdcdc")
        log_scroll = ttk.Scrollbar(logf, orient=VERTICAL, command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

        # 底部操作
        bottom = ttk.Frame(outer)
        bottom.pack(fill=X, pady=(10, 0))
        self.start_btn = ttk.Button(bottom, text="开始转换", command=self.start)
        self.start_btn.pack(side=LEFT)
        self.cancel_btn = ttk.Button(bottom, text="停止", command=self.cancel, state="disabled")
        self.cancel_btn.pack(side=LEFT, padx=6)
        self.open_btn = ttk.Button(bottom, text="打开输出目录", command=self.open_outdir, state="disabled")
        self.open_btn.pack(side=LEFT)
        ttk.Button(bottom, text="退出", command=self._on_close).pack(side=RIGHT)

    def _bind_dnd(self) -> None:
        if _DND_OK and hasattr(self.listbox, "drop_target_register"):
            try:
                self.listbox.drop_target_register(DND_FILES)
                self.listbox.dnd_bind("<<Drop>>", self._on_drop)
                self._log("已启用拖拽：可直接把文件拖到列表里。")
                return
            except Exception as exc:  # noqa: BLE001
                self._log(f"拖拽初始化失败：{exc}")
        self._log("拖拽支持不可用（缺少 tkinterdnd2），请使用「添加文件」按钮。")

    def _on_drop(self, event) -> None:
        try:
            items = self.root.tk.splitlist(event.data)
        except Exception:  # noqa: BLE001
            items = str(event.data).split()
        self._add_paths(items)

    # ------------------------------------------------------------ 引擎信息

    def _refresh_engine_label(self, force: bool = True) -> None:
        """重新探测引擎并更新顶部状态标签。"""
        if force:
            converter.clear_engine_cache()   # 关键：清缓存，否则「重新检测」是无效按钮
        desc = converter.describe_engine()
        ok = converter.word_engine_available() or bool(converter.com_engine_progid())
        self.engine_label.configure(
            text=f"Word → PDF：{desc}",
            foreground="#0a7a35" if ok else "#c0392b",
        )

    def _show_diagnostics(self) -> None:
        """弹出完整诊断报告，便于定位「未检测到 Word」类问题。"""
        converter.clear_engine_cache()
        try:
            report = converter.diagnose_engine()
        except Exception:  # noqa: BLE001
            report = "诊断过程出错：\n" + traceback.format_exc()

        win = tk.Toplevel(self.root)
        win.title("诊断报告")
        win.geometry("760x560")
        frame = ttk.Frame(win, padding=8)
        frame.pack(fill=BOTH, expand=True)
        text = tk.Text(frame, wrap="word", font=("Consolas", 9),
                       background="#1e1e1e", foreground="#dcdcdc")
        scroll = ttk.Scrollbar(frame, orient=VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)
        text.insert(END, report)
        text.configure(state="disabled")

        bar = ttk.Frame(win)
        bar.pack(fill=X, padx=8, pady=(0, 8))

        def copy_report() -> None:
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
            self._log("诊断报告已复制到剪贴板。")

        def save_report() -> None:
            path = filedialog.asksaveasfilename(
                title="保存诊断报告", defaultextension=".txt",
                initialfile="引擎诊断报告.txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(report)
                self._log(f"诊断报告已保存到：{path}")
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("保存失败", str(exc))

        ttk.Button(bar, text="复制到剪贴板", command=copy_report).pack(side=LEFT)
        ttk.Button(bar, text="保存为文件", command=save_report).pack(side=LEFT, padx=6)
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side=RIGHT)
        self._log("已生成诊断报告。")

    # ------------------------------------------------------------ 文件管理

    def _add_paths(self, paths) -> None:
        added = 0
        skipped = 0
        for path in paths:
            path = os.path.abspath(path)
            if os.path.isdir(path):
                for name in sorted(os.listdir(path)):
                    full = os.path.join(path, name)
                    if os.path.isfile(full) and converter.file_kind(full) != "unknown":
                        if self._push(full):
                            added += 1
                continue
            if not os.path.isfile(path):
                skipped += 1
                continue
            if converter.file_kind(path) == "unknown":
                skipped += 1
                continue
            if self._push(path):
                added += 1
            else:
                skipped += 1
        self._update_counts()
        self._log(f"已添加 {added} 个文件" + (f"，忽略 {skipped} 个不支持的项目" if skipped else ""))

    def _push(self, path: str) -> bool:
        if path in self.files:
            return False
        self.files.append(path)
        self.listbox.insert(END, self._label_for(path))
        return True

    def _label_for(self, path: str) -> str:
        kind = converter.file_kind(path)
        tag = "PDF" if kind == "pdf" else "Word"
        return f"[{tag}]  {os.path.basename(path)}    —  {os.path.dirname(path)}"

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[
                ("所有支持的文件", "*.docx *.doc *.docm *.rtf *.odt *.txt *.pdf"),
                ("Word 文档", "*.docx *.doc *.docm *.rtf *.odt *.txt"),
                ("PDF 文件", "*.pdf"),
                ("所有文件", "*.*"),
            ],
        )
        if paths:
            self._add_paths(paths)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含待转换文件的文件夹")
        if folder:
            self._add_paths([folder])

    def remove_selected(self) -> None:
        for index in sorted(self.listbox.curselection(), reverse=True):
            self.listbox.delete(index)
            del self.files[index]
        self._update_counts()

    def clear_files(self) -> None:
        self.listbox.delete(0, END)
        self.files.clear()
        self._update_counts()

    def _update_counts(self) -> None:
        w = sum(1 for f in self.files if converter.file_kind(f) == "word")
        p = sum(1 for f in self.files if converter.file_kind(f) == "pdf")
        mode = self.direction.get()
        if mode == "word2pdf":
            n = w
        elif mode == "pdf2word":
            n = p
        else:
            n = w + p
        self.status.configure(text=f"就绪：共 {len(self.files)} 个文件（Word {w} / PDF {p}），本次将转换 {n} 个")

    # ------------------------------------------------------------ 输出设置

    def _toggle_outdir(self) -> None:
        state = "disabled" if self.same_as_source.get() else "normal"
        self.out_entry.configure(state=state)

    def choose_outdir(self) -> None:
        folder = filedialog.askdirectory(title="选择输出目录", initialdir=self.outdir.get() or None)
        if folder:
            self.outdir.set(folder)
            self.same_as_source.set(False)
            self._toggle_outdir()

    def open_outdir(self) -> None:
        target = self.outdir.get()
        if os.path.isdir(target):
            try:
                os.startfile(target)  # type: ignore[attr-defined]
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("打不开目录", str(exc))

    # ------------------------------------------------------------ 日志

    def _log(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.insert(END, f"[{stamp}] {text}\n")
        self.log.see(END)

    # ------------------------------------------------------------ 转换流程

    def _plan_jobs(self) -> list[converter.Job]:
        mode = self.direction.get()
        jobs: list[converter.Job] = []
        for src in self.files:
            kind = converter.file_kind(src)
            if mode == "word2pdf":
                direction = "word2pdf"
                if kind != "word":
                    continue
            elif mode == "pdf2word":
                direction = "pdf2word"
                if kind != "pdf":
                    continue
            else:
                direction = "word2pdf" if kind == "word" else "pdf2word"
            if self.same_as_source.get():
                outdir = os.path.dirname(src)
            else:
                outdir = self.outdir.get().strip()
            if not outdir:
                outdir = os.path.dirname(src)
            dst = converter.default_output(src, outdir, direction)
            jobs.append(converter.Job(src=src, dst=dst, direction=direction))
        return jobs

    def start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        jobs = self._plan_jobs()
        if not jobs:
            messagebox.showinfo(
                "没有可转换的文件",
                "当前方向下没有匹配的文件。\n"
                "Word → PDF 需要 .doc/.docx 等文件；PDF → Word 需要 .pdf 文件。",
            )
            return

        # 输出目录可用性检查
        if not self.same_as_source.get():
            target = self.outdir.get().strip()
            try:
                os.makedirs(target, exist_ok=True)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("输出目录不可用", f"{target}\n\n{exc}")
                return

        self.cancel_flag.clear()
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.open_btn.configure(state="disabled")
        self.progress.configure(maximum=len(jobs), value=0)
        self.log.delete("1.0", END)
        self._log(f"开始转换，共 {len(jobs)} 个任务")

        options = converter.PdfToWordOptions(
            detect_headings=self.opt_headings.get(),
            keep_images=self.opt_images.get(),
            keep_tables=self.opt_tables.get(),
            page_break=self.opt_pagebreak.get(),
            scanned_fallback=self.opt_scanned.get(),
        )
        self.worker = threading.Thread(
            target=self._run, args=(jobs, options), daemon=True
        )
        self.worker.start()

    def _run(self, jobs, options) -> None:
        """在后台线程执行转换，通过队列与界面通信。"""
        done = 0

        def progress(text: str) -> None:
            self.msg_queue.put(("status", text))

        try:
            outcomes = converter.run_jobs(
                jobs,
                options=options,
                progress=progress,
                cancel=self.cancel_flag.is_set,
            )
            for job, error in outcomes:
                done += 1
                name = os.path.basename(job.src)
                if error:
                    tag = "取消" if error == "已取消" else "失败"
                    self.msg_queue.put(("log", f"✗ {tag}：{name} —— {error}"))
                else:
                    self.msg_queue.put(("log", f"✓ 完成：{name}  →  {os.path.basename(job.dst)}"))
                self.msg_queue.put(("progress", done))
            ok = sum(1 for _, e in outcomes if e is None)
            bad = len(outcomes) - ok
            self.msg_queue.put(("done", (ok, bad, bool(self.cancel_flag.is_set()))))
        except Exception:  # noqa: BLE001
            self.msg_queue.put(("log", "内部错误：\n" + traceback.format_exc()))
            self.msg_queue.put(("done", (done, len(jobs) - done, False)))

    def _pump(self) -> None:
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "status":
                    self.status.configure(text=str(payload))
                elif kind == "progress":
                    self.progress.configure(value=int(payload))
                elif kind == "log":
                    self._log(str(payload))
                elif kind == "done":
                    ok, bad, cancelled = payload  # type: ignore[misc]
                    self.start_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")
                    self.open_btn.configure(state="normal")
                    if cancelled:
                        self.status.configure(text=f"已停止：成功 {ok} 个，失败/跳过 {bad} 个")
                        self._log(f"—— 用户已停止。成功 {ok} 个，失败/跳过 {bad} 个 ——")
                    else:
                        self.status.configure(text=f"全部结束：成功 {ok} 个，失败 {bad} 个")
                        self._log(f"—— 转换结束。成功 {ok} 个，失败 {bad} 个 ——")
                    if ok and not bad:
                        try:
                            self.root.bell()
                        except Exception:  # noqa: BLE001
                            pass
        except queue.Empty:
            pass
        self.root.after(80, self._pump)

    def cancel(self) -> None:
        self.cancel_flag.set()
        self.status.configure(text="正在停止…（当前文件完成后中断）")
        self._log("已请求停止。")

    # ------------------------------------------------------------ 帮助

    def _show_help(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("转换说明")
        win.geometry("620x460")
        text = tk.Text(win, wrap="word", font=("Microsoft YaHei UI", 10), padx=14, pady=12)
        text.pack(fill=BOTH, expand=True)
        text.insert(END, HELP_TEXT)
        text.configure(state="disabled")
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=8)


HELP_TEXT = """Word → PDF
    调用本机 Microsoft Word 排版后导出，效果与在 Word 里「另存为 PDF」完全一致，
    文字可选中、可搜索，目录书签也会保留。首次调用需要启动 Word，约 10 秒；之后批量转换会快很多。

PDF → Word
    使用 PyMuPDF 解析 PDF 版面，再由 python-docx 生成可编辑的 .docx：
      · 按字号、加粗、编号样式推断标题层级
      · 还原表格为真正的 Word 表格
      · 提取页面图片并按位置插入
      · 扫描件（没有文字层）自动整页转成图片，保证内容不丢

    注意：PDF 格式本身不保存「段落」信息，任何工具都无法 100% 还原原排版。
    本工具的目标是「内容完整、结构可编辑」，复杂分栏或公式可能出现顺序偏差。

常见问题
    · 提示找不到 Word：需要安装 Microsoft Word（推荐）或 LibreOffice。
    · Word 转换失败提示「未能引发事件」：先手动打开一次 Word 并关闭，
      确认没有弹窗（如登录、激活、更新提示）后再重试。
    · 转换很慢：首次启动 Word 较慢属正常；扫描件转图片会明显变慢。
    · 文件被占用：请先关闭正在用 Word/PDF 阅读器打开的同名文件。
"""


def _setup_paths() -> None:
    """把程序自带的依赖目录加入搜索路径（兼容直接双击运行）。"""
    for sub in ("Lib\\site-packages", "lib\\site-packages"):
        candidate = os.path.join(HERE, "env", sub)
        if os.path.isdir(candidate) and candidate not in sys.path:
            sys.path.insert(0, candidate)


def main() -> int:
    _setup_paths()
    try:
        if _DND_OK:
            root = TkinterDnD.Tk()
        else:
            root = tk.Tk()
    except Exception:  # noqa: BLE001
        root = tk.Tk()
    try:
        ConverterApp(root)
    except Exception:  # noqa: BLE001
        messagebox.showerror("启动失败", traceback.format_exc())
        return 1
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
