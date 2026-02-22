from __future__ import annotations

import json
import os
import sys

import win32com.client  # type: ignore
import win32gui  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def invoke_read(jv):
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    dataspec = os.environ.get("JV_DATASPEC", "RACE")
    fromdate = os.environ.get("JV_FROMDATE", "20260222000000")
    option = int(os.environ.get("JV_OPTION", "1"))

    hwnd = win32gui.GetForegroundWindow()

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.ParentHWnd = hwnd  # 親ウィンドウ設定（UIが裏に行かないように）

    init_ret = jv.JVInit(0)
    ui_ret = jv.JVSetUIProperties()

    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)

    open_t = invoke_open(jv, dataspec, fromdate, option)
    open_ret, readcount, downloadcount, lastts = open_t

    read_ret, buff, size, filename = invoke_read(jv)

    jv.JVClose()

    payload = {
        "ok": True,
        "hwnd": int(hwnd),
        "setup": {"init": int(init_ret), "ui": int(ui_ret)},
        "open": {
            "ret": int(open_ret),
            "readcount": int(readcount),
            "downloadcount": int(downloadcount),
            "lastfiletimestamp": str(lastts),
        },
        "read": {"ret": int(read_ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:200]},
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())