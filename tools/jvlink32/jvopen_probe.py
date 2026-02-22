from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def _try(desc: str, fn):
    try:
        r = fn()
        return {"ok": True, "call": desc, "result": str(r)}
    except Exception as e:
        return {"ok": False, "call": desc, "error": str(e)}


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    init = _try("JVInit(0)", lambda: jv.JVInit(0))
    set_ui = _try("JVSetUIProperties()", lambda: jv.JVSetUIProperties())

    # 設定画面と同じ保存先
    savepath = r"C:\ProgramData\JRA-VAN\Data"
    set_savepath = _try(f"JVSetSavePath({savepath!r})", lambda: jv.JVSetSavePath(savepath))
    set_saveflag = _try("JVSetSaveFlag(1)", lambda: jv.JVSetSaveFlag(1))

    ds = "RACE"
    key1 = "20260222"
    key2 = "20260331"

    open_candidates = [
        _try("JVOpen(ds)", lambda: jv.JVOpen(ds)),
        _try("JVOpen(ds, key1)", lambda: jv.JVOpen(ds, key1)),
        _try("JVOpen(ds, key1, key2)", lambda: jv.JVOpen(ds, key1, key2)),
        _try("JVOpen(ds, key1, 0)", lambda: jv.JVOpen(ds, key1, 0)),
        _try("JVOpen(ds, key1, 1)", lambda: jv.JVOpen(ds, key1, 1)),
        _try("JVOpen(ds, key1, '0')", lambda: jv.JVOpen(ds, key1, "0")),
        _try("JVOpen(ds, key1, '1')", lambda: jv.JVOpen(ds, key1, "1")),
        _try("JVOpen(ds, key1, key2, 0)", lambda: jv.JVOpen(ds, key1, key2, 0)),
        _try("JVOpen(ds, key1, key2, 1)", lambda: jv.JVOpen(ds, key1, key2, 1)),
        _try("JVOpen(ds, key1, key2, '0')", lambda: jv.JVOpen(ds, key1, key2, "0")),
        _try("JVOpen(ds, key1, key2, '1')", lambda: jv.JVOpen(ds, key1, key2, "1")),
    ]

    payload = {
        "ok": True,
        "init": init,
        "set_ui": set_ui,
        "set_savepath": set_savepath,
        "set_saveflag": set_saveflag,
        "open_candidates": open_candidates,
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())