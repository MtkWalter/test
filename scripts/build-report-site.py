"""Build a self-hosted static site; user workbooks never leave the browser."""
from __future__ import annotations

import hashlib
import importlib.metadata
import json
import shutil
import subprocess
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from excel_insights.monthly_report import ReportRules  # noqa: E402


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def add_file(archive, path, name):
    from zipfile import ZipInfo
    info = ZipInfo(str(name).replace("\\", "/"), (2020, 1, 1, 0, 0, 0))
    info.create_system = 3
    info.external_attr = 0o644 << 16
    archive.writestr(info, path.read_bytes(), compress_type=ZIP_DEFLATED)


def build(destination: Path = ROOT / "_site"):
    runtime = ROOT / "node_modules" / "pyodide"
    info = json.loads((runtime / "package.json").read_text())
    if info["version"] != "0.29.5":
        raise ValueError("Run npm ci: the static site requires Pyodide 0.29.5.")
    destination.mkdir(parents=True, exist_ok=True)
    for source in (ROOT / "src/excel_insights/web").iterdir():
        if source.suffix in {".html", ".js", ".css"}:
            shutil.copyfile(source, destination / source.name)
    for source in (ROOT / "src/excel_insights/site").iterdir():
        if source.suffix == ".js":
            shutil.copyfile(source, destination / source.name)
    runtime_dest = destination / "runtime"
    runtime_dest.mkdir(exist_ok=True)
    for name in ("pyodide.js", "pyodide.asm.js", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json", "package.json", "README.md"):
        shutil.copyfile(runtime / name, runtime_dest / name)
    with ZipFile(destination / "python-deps.zip", "w", ZIP_DEFLATED) as archive:
        for package, expected in (("openpyxl", "3.1.5"), ("et-xmlfile", "2.0.0")):
            distribution = importlib.metadata.distribution(package)
            if distribution.version != expected:
                raise ValueError(f"Expected {package} {expected}; run uv sync --locked.")
            for entry in sorted(distribution.files or [], key=str):
                name = str(entry).replace("\\", "/")
                if name.startswith("../") or "__pycache__" in name or name.endswith(".pyc") or Path(name).name in {"INSTALLER", "RECORD", "REQUESTED", "direct_url.json"}:
                    continue
                path = distribution.locate_file(entry)
                if path.is_file():
                    add_file(archive, path, name)
    with ZipFile(destination / "report-engine.zip", "w", ZIP_DEFLATED) as archive:
        for name in ("__init__.py", "monthly_report.py", "report_preview.py", "raw_export.py"):
            add_file(archive, ROOT / "src/excel_insights" / name, "excel_insights/" + name)
    revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    manifest = {"revision": revision, "runtime": "0.29.5", "rules": ReportRules().to_dict(),
                "dependencies_sha256": digest(destination / "python-deps.zip"),
                "engine_sha256": digest(destination / "report-engine.zip")}
    (destination / "site-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    html = (destination / "index.html").read_text(encoding="utf-8")
    html = html.replace('<body>', '<body data-report-mode="static">')
    html = html.replace('<script src="app.js" defer></script>', '<script src="site-bridge.js" defer></script>\n  <script src="app.js" defer></script>')
    html = html.replace('href="style.css"', f'href="style.css?v={revision}"')
    html = html.replace('href="report.css"', f'href="report.css?v={revision}"')
    for name in ("site-bridge.js", "app.js"):
        html = html.replace(f'src="{name}"', f'src="{name}?v={revision}"')
    html = html.replace('本機工作階段', '瀏覽器工作階段').replace('本機運算', '瀏覽器內運算')
    html = html.replace('本機處理', '檔案留在瀏覽器').replace('關閉程式', '清除資料')
    html = html.replace('本機程式', '瀏覽器').replace('本機', '瀏覽器')
    html = html.replace('檔案只傳送到瀏覽器，不會上傳到 GitHub 或外部伺服器。', 'Excel 只在瀏覽器記憶體中讀取與計算，不會上傳到 GitHub 或外部伺服器。')
    html = html.replace('用右上角「清除資料」結束服務並清除暫存；已下載的 Excel 會保留。', '用右上角「清除資料」釋放工作資料；已下載的 Excel 會保留。')
    (destination / "index.html").write_text(html, encoding="utf-8")
    (destination / ".nojekyll").touch()
    (destination / "NOTICE.txt").write_text("Monthly Report static client\nPyodide 0.29.5: https://github.com/pyodide/pyodide (MPL-2.0)\nPython: https://www.python.org/psf/license/\nopenpyxl 3.1.5 and et-xmlfile 2.0.0: licenses included in python-deps.zip\n", encoding="utf-8")
    print(f"Static site built: {destination}")


if __name__ == "__main__":
    build()
