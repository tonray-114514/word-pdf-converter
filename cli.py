"""命令行入口：无界面批量转换（也可被脚本调用）。

示例
----
    python cli.py 报告.docx
    python cli.py --to-pdf a.docx b.doc -o D:\\out
    python cli.py --to-word 手册.pdf -o D:\\out --no-images
    python cli.py D:\\文档目录 -o D:\\out
"""

from __future__ import annotations

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


def _bootstrap_env() -> None:
    """确保使用程序自带的 env 运行环境。

    被系统 Python 直接启动时（例如 python cli.py --diagnose），
    自动换成自带 env 重新启动，避免因系统 Python 缺少 pywin32
    而误报「未检测到 Microsoft Word」。
    """
    if os.environ.get("WPC_REEXEC") == "1":
        return
    env_dir = os.path.join(HERE, "env")
    try:
        if os.path.normcase(os.path.abspath(sys.prefix)) == os.path.normcase(env_dir):
            return
    except Exception:  # noqa: BLE001
        pass
    if any("site-packages" in p and os.path.normcase(env_dir) in os.path.normcase(p)
           for p in sys.path):
        return

    venv_py = os.path.join(env_dir, "Scripts", "python.exe")
    if not os.path.isfile(venv_py):
        return
    # 自带环境里必须有 pywin32（pywin32_system32 下有 DLL），否则不切换
    site = os.path.join(env_dir, "Lib", "site-packages")
    system32 = os.path.join(site, "pywin32_system32")
    if not os.path.isdir(system32):
        return
    if not any(name.lower().startswith("pywintypes") and name.lower().endswith(".dll")
               for name in os.listdir(system32)):
        return

    os.environ["WPC_REEXEC"] = "1"
    try:
        os.execv(venv_py, [venv_py, os.path.abspath(__file__), *sys.argv[1:]])
    except Exception:  # noqa: BLE001
        os.environ.pop("WPC_REEXEC", None)


_bootstrap_env()

import converter  # noqa: E402


def collect(paths: list[str]) -> list[str]:
    files: list[str] = []
    for path in paths:
        if os.path.isdir(path):
            for name in sorted(os.listdir(path)):
                full = os.path.join(path, name)
                if os.path.isfile(full) and converter.file_kind(full) != "unknown":
                    files.append(full)
        elif os.path.isfile(path):
            files.append(path)
        else:
            print(f"跳过（不存在）：{path}", file=sys.stderr)
    return files


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="cli.py",
        description="Word ⇄ PDF 批量互转（命令行）",
    )
    parser.add_argument("inputs", nargs="+", help="文件或文件夹，可写多个")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--to-pdf", action="store_true", help="强制 Word → PDF")
    group.add_argument("--to-word", action="store_true", help="强制 PDF → Word")
    parser.add_argument("-o", "--outdir", default=None,
                        help="输出目录（默认与源文件同目录）")
    parser.add_argument("--same-dir", action="store_true", help="输出到源文件所在目录")
    parser.add_argument("--no-images", action="store_true", help="PDF→Word 时不提取图片")
    parser.add_argument("--no-tables", action="store_true", help="PDF→Word 时不还原表格")
    parser.add_argument("--no-headings", action="store_true", help="PDF→Word 时不识别标题")
    parser.add_argument("--no-pagebreak", action="store_true", help="PDF→Word 时不按页分页")
    parser.add_argument("--list-engine", action="store_true", help="只显示引擎信息后退出")
    parser.add_argument("--diagnose", action="store_true",
                        help="输出完整诊断报告后退出（排查「未检测到 Word」用）")
    args = parser.parse_args(argv)

    print(f"引擎：{converter.describe_engine()}")
    print("PyMuPDF / python-docx：", end="")
    try:
        import docx  # noqa: F401
        import pymupdf

        print(f"OK（PyMuPDF {pymupdf.__doc__ or pymupdf.version}）")
    except Exception as exc:  # noqa: BLE001
        print(f"缺失（{exc}）")
    if args.diagnose:
        print()
        print(converter.diagnose_engine())
        return 0
    if args.list_engine:
        return 0

    sources = collect(args.inputs)
    if not sources:
        print("没有找到可转换的文件。", file=sys.stderr)
        return 2

    jobs: list[converter.Job] = []
    for src in sources:
        kind = converter.file_kind(src)
        if args.to_pdf:
            direction = "word2pdf"
            if kind != "word":
                print(f"跳过（不是 Word 文件）：{src}", file=sys.stderr)
                continue
        elif args.to_word:
            direction = "pdf2word"
            if kind != "pdf":
                print(f"跳过（不是 PDF）：{src}", file=sys.stderr)
                continue
        else:
            if kind == "unknown":
                print(f"跳过（无法识别的格式）：{src}", file=sys.stderr)
                continue
            direction = "word2pdf" if kind == "word" else "pdf2word"
        outdir = os.path.dirname(src) if (args.same_dir or not args.outdir) else args.outdir
        jobs.append(converter.Job(
            src=os.path.abspath(src),
            dst=converter.default_output(src, outdir, direction),
            direction=direction,
        ))

    if not jobs:
        print("没有需要转换的任务。", file=sys.stderr)
        return 2

    options = converter.PdfToWordOptions(
        detect_headings=not args.no_headings,
        keep_images=not args.no_images,
        keep_tables=not args.no_tables,
        page_break=not args.no_pagebreak,
    )

    print(f"共 {len(jobs)} 个任务，开始转换…\n")
    outcomes = converter.run_jobs(
        jobs,
        options=options,
        progress=lambda text: print(f"  {text}", flush=True),
    )

    ok = 0
    for job, error in outcomes:
        name = os.path.basename(job.src)
        if error:
            print(f"✗ {name}：{error}")
        else:
            ok += 1
            print(f"✓ {name}  →  {job.dst}")
    print(f"\n完成：成功 {ok} 个，失败 {len(outcomes) - ok} 个。")
    return 0 if ok == len(outcomes) else 1


if __name__ == "__main__":
    try:
        code = main()
    finally:
        try:
            converter.cleanup_temp()
        except Exception:  # noqa: BLE001
            pass
    sys.exit(code)
