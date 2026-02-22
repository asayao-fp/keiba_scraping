from __future__ import annotations

import json
import sys


def main() -> int:
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        print(json.dumps({"ok": False, "step": "import_pywin32", "error": str(e)}), file=sys.stdout)
        return 2

    try:
        win32com.client.Dispatch("JVDTLab.JVLink")
        print(json.dumps({"ok": True, "step": "dispatch", "prog_id": "JVDTLab.JVLink"}), file=sys.stdout)
        return 0
    except Exception as e:
        print(json.dumps({"ok": False, "step": "dispatch", "prog_id": "JVDTLab.JVLink", "error": str(e)}), file=sys.stdout)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())