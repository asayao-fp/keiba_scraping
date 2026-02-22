from __future__ import annotations

import json
import os
import sys

import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    # JVOpen returns tuple like: (ret, readcount, downloadcount, lastfiletimestamp)
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def invoke_read(jv):
    # JVRead returns tuple like: (ret, buff, size, filename)
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    dataspec = os.environ.get("JV_DATASPEC", "RA")
    fromdate = os.environ.get("JV_FROMDATE", "20260222")
    option = int(os.environ.get("JV_OPTION", "0"))

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    init_ret = jv.JVInit(0)
    ui_ret = jv.JVSetUIProperties()
    savepath_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    saveflag_ret = jv.JVSetSaveFlag(1)

    st0 = jv.JVStatus()

    open_t = invoke_open(jv, dataspec, fromdate, option)
    # open_t is expected: (ret, readcount, downloadcount, lastfiletimestamp)
    open_ret, readcount, downloadcount, lastts = open_t

    st1 = jv.JVStatus()

    read_t = invoke_read(jv)
    read_ret, buff, size, filename = read_t

    st2 = jv.JVStatus()

    close_ret = jv.JVClose()

    payload = {
        "ok": True,
        "setup": {
            "init": int(init_ret),
            "ui": int(ui_ret),
            "savepath": int(savepath_ret),
            "saveflag": int(saveflag_ret),
        },
        "status": {"before_open": int(st0), "after_open": int(st1), "after_read": int(st2)},
        "open": {
            "dataspec": dataspec,
            "fromdate": fromdate,
            "option": option,
            "ret": int(open_ret),
            "readcount": int(readcount),
            "downloadcount": int(downloadcount),
            "lastfiletimestamp": str(lastts),
        },
        "read": {
            "ret": int(read_ret),
            "size": int(size),
            "filename": str(filename),
            "buff_head": str(buff)[:200],
        },
        "close": int(close_ret),
    }

    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())