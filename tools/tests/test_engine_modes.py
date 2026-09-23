"""验证 can_word2pdf 按引擎模式判断（而不是「本机有没有某个引擎」）。

用法: python test_engine_modes.py
"""
import os
import sys
import threading

APP = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, APP)
sys.path.insert(0, os.path.join(APP, "env", "Lib", "site-packages"))

import converter  # noqa: E402
import server  # noqa: E402

results = []


def check(name, cond, extra=""):
    print(("通过  " if cond else "失败  ") + name + (f"  {extra}" if extra else ""))
    results.append(bool(cond))


def make_manager(engine):
    mgr = server.JobManager.__new__(server.JobManager)   # 不启动后台线程
    mgr.jobs = {}
    mgr.lock = threading.Lock()
    mgr.engine = engine
    return mgr


has_word = converter.word_engine_available() or bool(converter.com_engine_progid())
has_lo = bool(converter.find_soffice())
print(f"本机情况: Word/WPS={has_word}  LibreOffice={has_lo}")
print()

print("=== auto 模式 ===")
st = make_manager("auto").stats()
check("can_word2pdf = 有任一引擎", st["can_word2pdf"] == (has_word or has_lo),
      f"实际 {st['can_word2pdf']}")

print("\n=== 强制 office 模式 ===")
st = make_manager("office").stats()
check("can_word2pdf = 有 Word/WPS", st["can_word2pdf"] == has_word,
      f"实际 {st['can_word2pdf']}")

print("\n=== 强制 libreoffice 模式 ===")
st = make_manager("libreoffice").stats()
check("can_word2pdf = 有 LibreOffice（关键修正）", st["can_word2pdf"] == has_lo,
      f"实际 {st['can_word2pdf']}")
if has_word and not has_lo:
    check("本机有 Word 但强制 LibreOffice 时应为 False",
          st["can_word2pdf"] is False, "正确识别出不可用")

print("\n=== 结果 ===")
print(f"通过 {sum(results)}/{len(results)}")
sys.exit(0 if all(results) else 1)
