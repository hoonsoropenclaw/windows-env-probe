import os
import pathlib
import subprocess
import sys
import tempfile

print("=== 1. stdout encoding + Chinese print ===")
print("stdout_encoding:", sys.stdout.encoding)
try:
    print("繁體中文測試：合併檔")
    print("CHINESE_STDOUT: OK")
except UnicodeEncodeError as e:
    print("CHINESE_STDOUT: FAIL UnicodeEncodeError:", e)

print("=== 2. reconfigure then print ===")
try:
    sys.stdout.reconfigure(encoding="utf-8")
    print("繁體中文測試：合併檔")
    print("CHINESE_AFTER_RECONFIGURE: OK")
except Exception as e:
    print("CHINESE_AFTER_RECONFIGURE: FAIL", type(e).__name__, e)

print("=== 3. openpyxl stack ===")
try:
    import openpyxl

    print("openpyxl:", openpyxl.__version__)
    d = pathlib.Path(tempfile.mkdtemp())
    p = d / "測試資料_2026.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明細"
    ws.append(["項目", "金額"])
    ws.append(["器材費", 120000])
    wb.save(p)
    wb2 = openpyxl.load_workbook(p)
    ws2 = wb2["明細"]
    print("OPENPYXL_ROUNDTRIP:", ws2.cell(row=2, column=1).value, ws2.cell(row=2, column=2).value)
    print("chinese_filename_exists:", p.exists())
except Exception as e:
    print("OPENPYXL: FAIL", type(e).__name__, e)

print("=== 4. cp950 filename / atomic replace ===")
try:
    d = pathlib.Path(tempfile.mkdtemp())
    src = d / "來源.xlsx"
    src.write_bytes(b"x")
    dst = d / "備份_中文名.xlsx"
    os.replace(src, dst)
    print("ATOMIC_REPLACE_CHINESE:", dst.exists())
except Exception as e:
    print("ATOMIC_REPLACE_CHINESE: FAIL", type(e).__name__, e)

print("=== 5. py launcher ===")
try:
    r = subprocess.run(["py", "-3", "-c", "import sys;print(sys.version.split()[0])"],
                       capture_output=True, text=True, timeout=60)
    print("PY_LAUNCHER:", r.returncode, (r.stdout or "").strip()[:40], (r.stderr or "").strip()[:80])
except Exception as e:
    print("PY_LAUNCHER: FAIL", type(e).__name__, e)

print("=== 6. pathlib separators ===")
print("PureWindowsPath parts:", pathlib.PureWindowsPath("C:/Users/x/Desktop/資料.xlsx").parts)
