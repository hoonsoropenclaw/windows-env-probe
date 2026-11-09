import os
import pathlib
import subprocess
import sys

print("=== A. Excel / COM availability ===")
try:
    import win32com.client  # noqa

    print("pywin32: INSTALLED")
except ImportError:
    print("pywin32: NOT INSTALLED")

try:
    import winreg

    found = []
    for root in (r"C:\Program Files\Microsoft Office", r"C:\Program Files (x86)\Microsoft Office"):
        p = pathlib.Path(root)
        if p.exists():
            found.append(str(p))
            for exe in p.rglob("EXCEL.EXE"):
                found.append("EXCEL:" + str(exe))
                break
    print("OFFICE_INSTALL:", found if found else "NOT FOUND")
except Exception as e:
    print("OFFICE_INSTALL: probe error", type(e).__name__, e)

print("=== B. UTF-8 mode ===")
print("PYTHONUTF8 env:", os.environ.get("PYTHONUTF8"))
print("utf8_mode attr:", getattr(sys.flags, "utf8_mode", None))
try:
    r = subprocess.run(
        [sys.executable, "-X", "utf8", "-c", "import sys;print(sys.stdout.encoding);print('繁體中文')"],
        capture_output=True,
        text=True,
        timeout=60,
    )
    print("PY_X_UTF8:", r.returncode, repr((r.stdout or "").strip()[:60]), repr((r.stderr or "").strip()[:80]))
except Exception as e:
    print("PY_X_UTF8: FAIL", type(e).__name__, e)

print("=== C. pandas ===")
try:
    import pandas as pd

    print("pandas:", pd.__version__)
except Exception as e:
    print("pandas: FAIL", type(e).__name__, e)

print("=== D. locale / chcp ===")
try:
    r = subprocess.run(["cmd", "/c", "chcp"], capture_output=True, text=True, timeout=30)
    print("CHCP:", (r.stdout or r.stderr or "").strip()[:60])
except Exception as e:
    print("CHCP: FAIL", type(e).__name__, e)

print("=== E. path semantics ===")
print("cwd:", os.getcwd())
print("sep:", os.sep, "| altsep:", os.altsep)
print("case_sensitive_os_probe:", pathlib.Path("A.txt").exists() == pathlib.Path("a.txt").exists())
print("reserved_name_ok:", end=" ")
try:
    d = pathlib.Path(os.environ.get("TEMP", "."))
    p = d / "CON.txt"
    print("CON.txt created:", p.exists())
except Exception as e:
    print("FAIL", type(e).__name__, e)
