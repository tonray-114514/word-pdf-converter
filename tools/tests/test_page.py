"""抽取页面里的 JS 并交给 Node 做语法检查（避免手工改动引入语法错误）。"""
import json
import os
import re
import subprocess
import sys
import tempfile

APP = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, APP)
sys.path.insert(0, os.path.join(APP, "env", "Lib", "site-packages"))

import server  # noqa: E402

html = server.PAGE_HTML

# 抓取最后一个 <script>...</script> 块（页面逻辑）
blocks = re.findall(r"<script>(.*?)</script>", html, re.S)
print(f"找到 {len(blocks)} 个内联 script 块")
if not blocks:
    sys.exit("未找到 script 块")

js = blocks[-1]
print(f"脚本长度: {len(js)} 字符")

# 基本结构核对
need = [
    "loadInfo", "applyEngineState", "renderFile", "poll", "applyStatus",
    "escapeHtml", "histbtn", "canWord2Pdf", "state.canWord2Pdf",
]
missing = [n for n in need if n not in js]
print("关键函数/变量:", "齐全" if not missing else f"缺失 {missing}")

# 括号配平（粗查）
for open_ch, close_ch in [("{", "}"), ("(", ")"), ("[", "]")]:
    a, b = js.count(open_ch), js.count(close_ch)
    print(f"  {open_ch}{close_ch} 配平: {a} vs {b} -> {'OK' if a == b else '不平衡'}")

# 交给 node 做真正的语法解析
node = None
for candidate in ("node", "node.exe"):
    from shutil import which

    node = which(candidate)
    if node:
        break
if not node:
    print("未找到 node，跳过语法解析")
    sys.exit(0)

tmp = os.path.join(tempfile.gettempdir(), "_jscheck.js")
with open(tmp, "w", encoding="utf-8") as fh:
    fh.write(js + "\n")
try:
    proc = subprocess.run([node, "--check", tmp], capture_output=True, text=True)
    if proc.returncode == 0:
        print("\nNode 语法检查: 通过")
    else:
        print("\nNode 语法检查: 失败")
        print(proc.stderr[:1500])
        sys.exit(1)
finally:
    try:
        os.remove(tmp)
    except OSError:
        pass

# 顺带核对 HTML 结构
for tag in ("html", "head", "body", "style"):
    opens = len(re.findall(rf"<{tag}[ >]", html))
    closes = len(re.findall(rf"</{tag}>", html))
    print(f"  <{tag}> {opens} 开 / {closes} 闭", "OK" if opens == closes else "不匹配")

print("\n界面资源自检通过")
