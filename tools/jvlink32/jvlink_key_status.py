from __future__ import annotations

import json
import os
import sys

import win32com.client  # type: ignore


def main() -> int:
    service_key = os.environ.get("JRAVAN_SERVICE_KEY", "")

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    init_ret = jv.JVInit(0)
    ui_ret = jv.JVSetUIProperties()
    savepath_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    saveflag_ret = jv.JVSetSaveFlag(1)

    st1 = jv.JVStatus()

    set_key_ret = None
    if service_key:
        set_key_ret = jv.JVSetServiceKey(service_key)

    st2 = jv.JVStatus()
    jv.JVClose()

    payload = {
        "ok": True,
        "init": str(init_ret),
        "ui": str(ui_ret),
        "savepath": str(savepath_ret),
        "saveflag": str(saveflag_ret),
        "status_before_key": str(st1),
        "set_service_key_ret": (str(set_key_ret) if set_key_ret is not None else None),
        "status_after_key": str(st2),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())