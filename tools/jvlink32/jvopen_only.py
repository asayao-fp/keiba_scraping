from __future__ import annotations

import json
import os
import sys

import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def main() -> int:
    dataspec = os.environ.get("JV_DATASPEC", "RACE")
    fromdate = os.environ.get("JV_FROMDATE", "20240101000000")
    option = int(os.environ.get("JV_OPTION", "1"))

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    init_ret = jv.JVInit(0)
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)
    jv.JVSetPayFlag(0)

    open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

    # Closeまでやって終了
    close_ret = jv.JVClose()

    payload = {
        "ok": True,
        "setup": {"init": int(init_ret)},
        "open": {
            "dataspec": dataspec,
            "fromdate": fromdate,
            "option": option,
            "ret": int(open_ret),
            "readcount": int(readcount),
            "downloadcount": int(downloadcount),
            "lastts": str(lastts),
        },
        "close": int(close_ret),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())