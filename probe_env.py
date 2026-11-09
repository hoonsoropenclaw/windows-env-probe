"""Windows 環境探測（office-automation admin 軌用）。

用途：在真實 Windows 上量測「本地推導」以外的事實，供 code SOP §5-4
「Windows 專屬行為的驗證揭露：實測平台／模擬平台／未驗證項目三分欄」使用。

事實來源：GitHub Actions `windows-latest`（windows-2025-vs2026, image 20261103.298）。
⚠️ 這是 Windows Server 2025 的 CI runner，**不是使用者的 Windows 桌機**：
   - 沒有安裝 Microsoft Office → COM 自動化仍屬「未驗證」
   - 桌面環境的 Excel 鎖檔／巨集／COM 行為不能由本 runner 代表
   但「Python 對 Windows 的行為」（codec／path／fs／atomic replace）可由本 runner 實測。
"""
import os
import pathlib
import subprocess
import sys
import tempfile

results = {}


def probe(name):
    def deco(fn):
        try:
            results[name] = fn()
        except Exception as e:  # noqa: BLE001
            results[name] = f"FAIL {type(e).__name__}: {e}"
        return fn

    return deco


print("=== 1. platform / python ===")
print("platform:", sys.platform)
print("python:", sys.version.split()[0])
print("fsencoding:", sys.getfilesystemencoding())
print("stdout_encoding:", sys.stdout.encoding)
print("utf8_mode:", getattr(sys.flags, "utf8_mode", None))
print("sep:", os.sep, "| altsep:", os.altsep)
print("cwd:", os.getcwd())

print("=== 2. stdout 中文（未 reconfigure）===")
try:
    print("繁體中文測試：合併檔")
    print("CHINESE_STDOUT: OK")
except UnicodeEncodeError as e:
    print("CHINESE_STDOUT: FAIL UnicodeEncodeError:", e)

print("=== 3. stdout 中文（reconfigure utf-8 後）===")
try:
    sys.stdout.reconfigure(encoding="utf-8")
    print("繁體中文測試：合併檔")
    print("CHINESE_AFTER_RECONFIGURE: OK")
except Exception as e:
    print("CHINESE_AFTER_RECONFIGURE: FAIL", type(e).__name__, e)

print("=== 4. cp950 / big5 編碼可用性 ===")
try:
    "繁體中文測試：合併檔".encode("cp950")
    print("cp950: OK")
except Exception as e:
    print("cp950: FAIL", e)

print("=== 5. openpyxl roundtrip（含中文檔名與工作表名）===")
try:
    import openpyxl

    d = pathlib.Path(tempfile.mkdtemp())
    p = d / "測試資料_2026.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明細"
    ws.append(["項目", "金額"])
    ws.append(["器材費", 120000])
    wb.save(p)
    ws2 = openpyxl.load_workbook(p)["明細"]
    print("openpyxl:", openpyxl.__version__)
    print("OPENPYXL_ROUNDTRIP:", ws2.cell(row=2, column=1).value, ws2.cell(row=2, column=2).value)
    print("chinese_filename_exists:", p.exists())
except Exception as e:
    print("OPENPYXL: FAIL", type(e).__name__, e)

print("=== 6. pandas ===")
try:
    import pandas as pd

    print("pandas:", pd.__version__)
except Exception as e:
    print("pandas: FAIL", type(e).__name__, e)

print("=== 7. 原子寫入（暫存檔 + os.replace）與中文檔名 ===")
try:
    d = pathlib.Path(tempfile.mkdtemp())
    src = d / "來源.xlsx"
    src.write_bytes(b"x")
    dst = d / "備份_中文名.xlsx"
    os.replace(src, dst)
    print("ATOMIC_REPLACE_CHINESE:", dst.exists())
except Exception as e:
    print("ATOMIC_REPLACE_CHINESE: FAIL", type(e).__name__, e)

print("=== 8. py launcher ===")
try:
    r = subprocess.run(
        ["py", "-3", "-c", "import sys;print(sys.version.split()[0])"],
        capture_output=True, text=True, timeout=60,
    )
    print("PY_LAUNCHER:", r.returncode, (r.stdout or "").strip()[:40])
except Exception as e:
    print("PY_LAUNCHER: FAIL", type(e).__name__, e)

print("=== 9. Microsoft Office / Excel COM 可用性 ===")
try:
    import win32com.client  # noqa: F401

    print("pywin32: INSTALLED")
except ImportError:
    print("pywin32: NOT INSTALLED")
found = []
for root in (r"C:\Program Files\Microsoft Office", r"C:\Program Files (x86)\Microsoft Office"):
    p = pathlib.Path(root)
    if p.exists():
        found.append(str(p))
        for exe in p.rglob("EXCEL.EXE"):
            found.append("EXCEL:" + str(exe))
            break
print("OFFICE_INSTALL:", found if found else "NOT FOUND")

print("=== 10. 保留裝置名稱 ===")
try:
    d = pathlib.Path(tempfile.mkdtemp())
    p = d / "CON.txt"
    p.write_text("x")
    print("CON.txt created:", p.exists())
except Exception as e:
    print("CON.txt: FAIL", type(e).__name__, e)

print("=== 11. 長路徑 ===")
try:
    d = pathlib.Path(tempfile.mkdtemp())
    p = d / ("x" * 250)
    p.write_text("x")
    print("long_path_write: OK")
except Exception as e:
    print("long_path_write: FAIL", type(e).__name__, e)

print("=== 12. chcp ===")
try:
    r = subprocess.run(["cmd", "/c", "chcp"], capture_output=True, text=True, timeout=30)
    print("CHCP:", (r.stdout or r.stderr or "").strip()[:60])
except Exception as e:
    print("CHCP: FAIL", type(e).__name__, e)

print("=== DONE ===")
