"""Build the portable Windows release: build/dist/ExcelToPDF/ and a versioned zip.

Usage (from the project root, inside the dev virtualenv):
    python tools/build_windows.py

Result: build/release/ExcelToPDF-Professional-<version>-win64.zip
Unzip anywhere and run ExcelToPDF.exe. Needs LibreOffice (free) or Microsoft Excel on the
target computer; the app offers to install LibreOffice when neither is found.
"""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from x2p import __version__  # noqa: E402

NAME = "ExcelToPDF"
BUILD = ROOT / "build"
DIST = BUILD / "dist"
RELEASE = BUILD / "release"
ENGINES = "src/x2p/engines"


def run_pyinstaller() -> Path:
    sep = ";" if sys.platform == "win32" else ":"
    cmd = [
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed",
        "--name", NAME, "--icon", str(ROOT / "packaging" / "app.ico"),
        "--paths", str(ROOT / "src"),
        "--distpath", str(DIST), "--workpath", str(BUILD / "work"), "--specpath", str(BUILD),
        # worker scripts run as files: the LibreOffice one by LibreOffice's own Python
        "--add-data", f"{ROOT / ENGINES / 'lo_worker.py'}{sep}x2p/engines",
        "--add-data", f"{ROOT / ENGINES / 'excel_worker.py'}{sep}x2p/engines",
        "--add-data", f"{ROOT / 'src/x2p/resources/app.png'}{sep}x2p/resources",
        "--hidden-import", "x2p.engines.excel_worker",
        "--hidden-import", "win32com.client", "--hidden-import", "pythoncom",
        "--exclude-module", "tkinter", "--exclude-module", "pytest",
        # unused heavy optional parts: numpy, PyMuPDF's ONNX layout models
        "--exclude-module", "numpy", "--exclude-module", "pymupdf.layout", "--exclude-module", "onnxruntime",
        str(ROOT / "packaging" / "x2p_app.py"),
    ]
    subprocess.run(cmd, check=True, cwd=ROOT)
    return DIST / NAME


UNNEEDED = [
    "_internal/PySide6/opengl32sw.dll",     # software OpenGL fallback; a widgets app does not need it
    "_internal/pymupdf/layout",             # ONNX models for PyMuPDF's AI layout analysis (unused)
]


def prune(app_dir: Path) -> None:
    for rel in UNNEEDED:
        p = app_dir / rel
        if p.is_dir():
            shutil.rmtree(p)
        elif p.is_file():
            p.unlink()


def add_docs(app_dir: Path) -> None:
    for f in ("README.md", "THIRD_PARTY_NOTICES.md", "CHANGELOG.md"):
        if (ROOT / f).is_file():
            shutil.copy2(ROOT / f, app_dir / f)
    (app_dir / "START-HERE.txt").write_text(  # ASCII name: safe in every unzip tool
        "Excel to PDF Professional\n\n"
        "1) شغّل ExcelToPDF.exe (لا يحتاج تثبيت Python).\n"
        "2) يحتاج البرنامج LibreOffice (مجاني) أو Microsoft Excel على الجهاز.\n"
        "   إن لم يجد أيًّا منهما سيعرض عليك تثبيت LibreOffice.\n"
        "3) اسحب ملف Excel إلى النافذة، راجع الأوراق، اضغط «معاينة PDF» ثم «تصدير».\n\n"
        "Run ExcelToPDF.exe. Requires LibreOffice (free) or Microsoft Excel.\n",
        encoding="utf-8")


def make_zip(app_dir: Path) -> Path:
    RELEASE.mkdir(parents=True, exist_ok=True)
    out = RELEASE / f"ExcelToPDF-Professional-{__version__}-win64.zip"
    out.unlink(missing_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for p in sorted(app_dir.rglob("*")):
            if p.is_file():
                z.write(p, Path(NAME) / p.relative_to(app_dir))
    digest = hashlib.sha256(out.read_bytes()).hexdigest()
    (out.parent / (out.name + ".sha256")).write_text(f"{digest}  {out.name}\n", encoding="ascii")
    return out


def main() -> int:
    app_dir = run_pyinstaller()
    prune(app_dir)
    add_docs(app_dir)
    z = make_zip(app_dir)
    print(f"app: {app_dir}\nzip: {z} ({z.stat().st_size / 1024**2:.1f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
