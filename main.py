"""PyInstaller entry point; capture startup failures even without a console."""
import json
from pathlib import Path
import sys
import traceback

if __name__ == "__main__":
    try:
        from excel_insights.app import main
        raise SystemExit(main())
    except Exception:
        details = traceback.format_exc()
        if "--smoke-output" in sys.argv:
            path = Path(sys.argv[sys.argv.index("--smoke-output") + 1])
        else:
            import tempfile
            path = Path(tempfile.gettempdir()) / "ExcelInsights-startup-error.json"
        path.write_text(json.dumps({"error": details}, ensure_ascii=False), encoding="utf-8")
        if "--smoke-test" not in sys.argv and sys.platform == "win32":
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, f"程式啟動失敗。詳細記錄：\n{path}", "Excel Insights", 16)
        raise SystemExit(1)
