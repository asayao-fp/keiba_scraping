from __future__ import annotations

import json
import os
import sys

import win32com.client  # type: ignore


def main() -> int:
    dataspec = os.environ.get("JV_DATASPEC", "RACE")
    fromdate = os.environ.get("JV_FROMDATE", "20260222")
    option = int(os.environ.get("JV_OPTION", "0"))

    service_key = os.environ.get("JRAVAN_SERVICE_KEY", "")

    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    init_ret = jv.JVInit(0)
    ui_ret = jv.JVSetUIProperties()
    savepath_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    saveflag_ret = jv.JVSetSaveFlag(1)

    set_key_ret = None
    if service_key:
        set_key_ret = jv.JVSetServiceKey(service_key)

    open_result = jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )

    read_result = jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )

    close_ret = jv.JVClose()

    payload = {
        "ok": True,
        "setup": {
            "init": str(init_ret),
            "ui": str(ui_ret),
            "savepath": str(savepath_ret),
            "saveflag": str(saveflag_ret),
            "set_service_key_called": bool(service_key),
            "set_service_key_ret": (str(set_key_ret) if set_key_ret is not None else None),
        },
        "open_invoke": str(open_result),
        "read_invoke": str(read_result),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())