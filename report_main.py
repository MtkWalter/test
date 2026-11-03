"""Separate PyInstaller entry point for MonthlyReport.exe."""
from pathlib import Path
import json
import sys
import tempfile
import traceback


if __name__ == "__main__":
    try:
        if "--input" in sys.argv:
            from excel_insights.monthly_report import main
        else:
            from excel_insights.monthly_app import main
        raise SystemExit(main())
    except Exception:
        details = traceback.format_exc()
        path = Path(tempfile.gettempdir()) / "MonthlyReport-startup-error.json"
        if "--smoke-output" in sys.argv:
            path = Path(sys.argv[sys.argv.index("--smoke-output") + 1])
        path.write_text(json.dumps({"error": details}, ensure_ascii=False), encoding="utf-8")
        if "--smoke-test" not in sys.argv and "--input" not in sys.argv and sys.platform == "win32":
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, f"程式啟動失敗。詳細記錄：\n{path}", "Monthly Report", 16)
        raise SystemExit(1)
