"""检查所有 .bat 文件的编码与换行符是否正确。

Windows 的 cmd.exe 有两个硬性要求，违反任何一个都会导致批处理执行错乱：
  1. 换行必须是 CRLF（LF-only 会让 cmd 把命令拆碎，报 'xxx' is not recognized）
  2. 中文内容必须用系统 ANSI 代码页（中文系统为 GBK/936）编码，
     UTF-8 中文会被按 GBK 误读，破坏命令与引号配对

本脚本用于守住这两点，避免以后重新生成批处理时再次踩坑。
"""
import os
import sys

APP = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ok = True
names = sorted(n for n in os.listdir(APP) if n.lower().endswith(".bat"))
if not names:
    print("未找到任何 .bat 文件")
    sys.exit(1)

print(f"检查 {len(names)} 个批处理文件")
print("=" * 64)

for name in names:
    path = os.path.join(APP, name)
    with open(path, "rb") as fh:
        data = fh.read()

    problems = []

    # 1) BOM 检查（BOM 会让 cmd 把第一行当命令）
    if data[:3] == b"\xef\xbb\xbf":
        problems.append("含 UTF-8 BOM")

    # 2) 换行检查
    cr = data.count(b"\r")
    lf = data.count(b"\n")
    if cr != lf:
        problems.append(f"换行不是 CRLF（CR={cr} LF={lf}）")
    if b"\n" in data.replace(b"\r\n", b""):
        # 存在落单的 LF
        lone = data.replace(b"\r\n", b"").count(b"\n")
        if lone:
            problems.append(f"存在 {lone} 个落单 LF")

    # 3) 编码检查：应能用 GBK 解码，且不应含明显的 UTF-8 多字节中文
    try:
        text = data.decode("gbk")
    except UnicodeDecodeError as exc:
        problems.append(f"GBK 解码失败：{exc}")
        text = ""

    # UTF-8 中文被 GBK 解码后会得到生僻字，抽常见的判断
    if text and any(ch in text for ch in "锟斤拷"):
        problems.append("疑似 UTF-8 被当作 GBK")

    # 4) 命令结构基础检查
    if text:
        first = text.splitlines()[0] if text.splitlines() else ""
        if not first.lower().startswith("@echo off") and not first.startswith("@"):
            problems.append(f"首行不是 @echo off（实际 {first[:20]!r}）")

    if problems:
        ok = False
        print(f"失败  {name}")
        for p in problems:
            print(f"        · {p}")
    else:
        print(f"通过  {name}  (CRLF, GBK, {len(data)} bytes)")

print("=" * 64)
print("全部批处理格式正确" if ok else "存在格式问题，双击可能无法正常启动")
sys.exit(0 if ok else 1)
