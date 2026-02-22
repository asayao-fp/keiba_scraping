from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def _get_attr(obj, name: str):
    try:
        return {"ok": True, "name": name, "value": getattr(obj, name)}
    except Exception as e:
        return {"ok": False, "name": name, "error": str(e)}


def _call(obj, name: str, *args):
    try:
        fn = getattr(obj, name)
        r = fn(*args)
        return {"ok": True, "name": name, "args": list(args), "result": str(r)}
    except Exception as e:
        return {"ok": False, "name": name, "args": list(args), "error": str(e)}


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    payload = {
        "ok": True,
        "jvlink_version": _get_attr(jv, "m_JVLinkVersion"),
        "savepath": _get_attr(jv, "m_savepath"),
        "saveflag": _get_attr(jv, "m_saveflag"),
        "servicekey": _get_attr(jv, "m_servicekey"),
        "payflag": _get_attr(jv, "m_payflag"),
        "status_before": _call(jv, "JVStatus"),
        "init": _call(jv, "JVInit", 0),
        "status_after_init": _call(jv, "JVStatus"),
        # ためしに UI 設定（親ウィンドウなし）
        "set_ui": _call(jv, "JVSetUIProperties", 0),
        "status_after_ui": _call(jv, "JVStatus"),
        # テストで RACE を open（あなたの結果再現）
        "open_race": _call(jv, "JVOpen", "RACE", "20260222", "20260331"),
        "read_once": _call(jv, "JVRead"),
        "close": _call(jv, "JVClose"),
        "status_after_close": _call(jv, "JVStatus"),
    }

    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())