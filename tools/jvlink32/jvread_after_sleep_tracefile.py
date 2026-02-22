from __future__ import annotations

import json
import os
import time
import traceback

import win32com.client  # type: ignore


LOG_PATH = os.path.join(os.path.dirname(__file__), "trace_jvread_sleep.log")


def log(obj):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps({"t": time.time(), **obj}, ensure_ascii=False) + "\n")
        f.flush()


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


def main():
    try:
        os.remove(LOG_PATH)
    except FileNotFoundError:
        pass

    dataspec = "RACE"
    fromdate = "20240101000000"
    option = 1

    sleep_sec = 20

    log({"step": "start", "sleep_sec": sleep_sec})

    try:
        jv = win32com.client.Dispatch("JVDTLab.JVLink")
        jv.JVInit(0)
        jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
        jv.JVSetSaveFlag(1)
        jv.JVSetPayFlag(0)

        log({"step": "JVOpen"})
        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)
        log({"step": "JVOpen_ret", "ret": int(open_ret), "readcount": int(readcount), "downloadcount": int(downloadcount), "lastts": str(lastts)})

        # 待機
        for i in range(sleep_sec):
            st = int(jv.JVStatus())
            log({"step": "sleep", "i": i + 1, "status": st})
            time.sleep(1)

        log({"step": "JVRead_once"})
        ret, buff, size, filename = invoke_read(jv)
        log({"step": "JVRead_ret", "ret": int(ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:80]})

        jv.JVClose()
        log({"step": "done"})
    except Exception:
        log({"step": "exception", "traceback": traceback.format_exc()})
        raise


if __name__ == "__main__":
    main()