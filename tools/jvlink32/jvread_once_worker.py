from __future__ import annotations

import gc
import json
import os
import sys
import time

import pythoncom  # type: ignore
import win32com.client  # type: ignore


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
    fromdate = os.environ.get("JV_FROMDATE", "20240101000000")
    option = int(os.environ.get("JV_OPTION", "1"))
    sleep_before_read = float(os.environ.get("JV_SLEEP_BEFORE_READ_SEC", "0"))

    pythoncom.CoInitialize()
    jv = None
    close_ret = None

    payload: dict = {"ok": False}
    try:
        jv = win32com.client.Dispatch("JVDTLab.JVLink")
        init_ret = jv.JVInit(0)
        savepath_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
        saveflag_ret = jv.JVSetSaveFlag(1)
        payflag_ret = jv.JVSetPayFlag(0)

        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

        if sleep_before_read > 0:
            time.sleep(sleep_before_read)

        # ここでクラッシュしても親は生きる
        read_ret, buff, size, filename = invoke_read(jv)

        payload = {
            "ok": True,
            "setup": {
                "init": int(init_ret),
                "savepath": int(savepath_ret),
                "saveflag": int(saveflag_ret),
                "payflag": int(payflag_ret),
            },
            "open": {
                "ret": int(open_ret),
                "readcount": int(readcount),
                "downloadcount": int(downloadcount),
                "lastts": str(lastts),
            },
            "read": {
                "ret": int(read_ret),
                "size": int(size),
                "filename": str(filename),
                "buff_head": str(buff)[:200],
            },
            "close": None,
        }
        return 0

    except Exception as e:
        payload = {"ok": False, "error": repr(e), "close": None}
        return 2

    finally:
        try:
            if jv is not None:
                close_ret = int(jv.JVClose())
        except Exception:
            close_ret = -9999

        jv = None
        gc.collect()
        pythoncom.CoUninitialize()

        if isinstance(payload, dict):
            payload["close"] = close_ret
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)


if __name__ == "__main__":
    raise SystemExit(main())