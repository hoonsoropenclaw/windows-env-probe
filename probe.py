import os
import pathlib
import platform
import shutil
import sys
import tempfile

print("platform:", platform.platform())
print("python:", sys.version.split()[0])
print("fsencoding:", sys.getfilesystemencoding())
print("stdout_encoding:", sys.stdout.encoding)

try:
    "繁體中文測試".encode("cp950")
    print("cp950: OK")
except Exception as e:
    print("cp950: FAIL", e)

try:
    d = pathlib.Path(tempfile.mkdtemp())
    f = d / "a.txt"
    f.write_text("x", encoding="utf-8")
    os.replace(f, d / "b.txt")
    print("os.replace:", "OK" if (d / "b.txt").exists() else "FAIL")
    print("max_path_usable:", len(str(d / ("x" * 200))))
except Exception as e:
    print("os.replace: FAIL", e)

print("has_py_launcher_alias:", shutil.which("py") is not None)
print("longpath_probe_len:", 300)
try:
    d = pathlib.Path(tempfile.mkdtemp())
    p = d / ("x" * 250)
    p.write_text("x")
    print("long_path_write:", "OK")
except Exception as e:
    print("long_path_write: FAIL", type(e).__name__, e)
