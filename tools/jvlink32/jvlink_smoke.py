from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta

import win32com.client  # type: ignore


def _safe_call(obj, name: str, *args):
    try:
        fn = getattr(obj, name)
        return {"ok": True, "name": name, "result": fn(*args)}
    except Exception as e:
        return {"ok": False, "name": name, "error": str(e)}


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    # 多くの実装で必要になるので、まずプロパティ/メソッドの存在を確認
    probe_names = [
        "JVInit",
        "JVOpen",
        "JVRead",
        "JVClose",
        "JVStatus",
        "JVSetUIProperties",
        "JVGetFileStatus",
        "JVGetNewData",
    ]
    exists = {n: hasattr(jv, n) for n in probe_names}

    # 直近日付（今日-7日）～今日+7日あたりを対象にしがちなので、その文字列も出しておく
    today = datetime.now().date()
    date_window = {
        "today": today.strftime("%Y%m%d"),
        "from": (today - timedelta(days=7)).strftime("%Y%m%d"),
        "to": (today + timedelta(days=7)).strftime("%Y%m%d"),
    }

    payload = {
        "ok": True,
        "prog_id": "JVDTLab.JVLink",
        "method_exists": exists,
        "date_window": date_window,
        "note": "This is a smoke test. Next step will choose the right DataSpec based on what is supported.",
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())