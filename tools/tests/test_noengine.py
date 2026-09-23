"""验证无引擎场景：can_word2pdf=false，且强制引擎时报错清晰。

不启动服务，直接测服务端的引擎判定逻辑。
"""
import os
import sys

APP = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, APP)
sys.path.insert(0, os.path.join(APP, "env", "Lib", "site-packages"))
SCRATCH = os.path.join(APP, ".tmp")
os.makedirs(SCRATCH, exist_ok=True)

import converter  # noqa: E402
import server  # noqa: E402

results = []


def check(name, cond, extra=""):
    print(("通过  " if cond else "失败  ") + name + (f"  {extra}" if extra else ""))
    results.append(bool(cond))


src = os.path.join(SCRATCH, "engine_probe.docx")
import docx  # noqa: E402

d = docx.Document()
d.add_paragraph("probe")
d.save(src)

dst = os.path.join(SCRATCH, "engine_probe.pdf")

print("=== 正常情况（本机有 Word）===")
check("word_engine_available", converter.word_engine_available() is True)
check("com_engine_progid", converter.com_engine_progid() == "Word.Application")

print("\n=== 模拟没有任何引擎 ===")
orig = (converter.find_word, converter.find_soffice, converter.com_engine_progid)
converter.find_word = lambda *a, **k: None
converter.find_soffice = lambda *a, **k: None
converter.com_engine_progid = lambda: None
converter.clear_engine_cache()
try:
    check("word_engine_available 变 false", converter.word_engine_available() is False)
    # stats 里的 can_word2pdf 应随之变 false
    mgr = server.JobManager.__new__(server.JobManager)   # 不启动线程
    mgr.jobs = {}
    mgr.lock = __import__("threading").Lock()
    mgr.engine = "auto"
    st = mgr.stats()
    check("stats.can_word2pdf=false", st["can_word2pdf"] is False, str(st))
    check("stats.engine 描述为无引擎", "无可用引擎" in st["engine"], st["engine"])

    # 此时转换应给出可诊断的错误
    try:
        converter.word_to_pdf(src, dst, engine="auto")
        check("无引擎时转换应失败", False, "竟然成功了")
    except converter.ConvertError as exc:
        msg = str(exc)
        check("报错包含排查指引", "诊断报告" in msg or "LibreOffice" in msg,
              msg.splitlines()[0])

    # 强制 libreoffice 但没有 LibreOffice -> 同样报错清晰
    try:
        converter.word_to_pdf(src, dst, engine="libreoffice")
        check("强制 libreoffice 应失败", False, "竟然成功了")
    except converter.ConvertError as exc:
        check("libreoffice 报错清晰", "LibreOffice" in str(exc))
finally:
    converter.find_word, converter.find_soffice, converter.com_engine_progid = orig
    converter.clear_engine_cache()
    for f in (src, dst):
        if os.path.exists(f):
            os.remove(f)

print("\n=== 恢复后 ===")
check("引擎判定已恢复", converter.word_engine_available() is True)
check("describe_engine 正常", "Word" in converter.describe_engine())

print()
print(f"通过 {sum(results)}/{len(results)}")
sys.exit(0 if all(results) else 1)
