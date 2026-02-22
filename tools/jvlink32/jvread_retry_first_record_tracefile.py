from __future__ import annotations

import json
import os
import time
import traceback

import win32com.client  # type: ignore


LOG_PATH = os.path.join(os.path.dirname(__file__), "trace_jvread_retry.log")


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
    # ログ初期化
    try:
        os.remove(LOG_PATH)
    except FileNotFoundError:
        pass

    dataspec = "RACE"
    fromdate = "20240101000000"
    option = 1

    log({"step": "start", "dataspec": dataspec, "fromdate": fromdate, "option": option})

    try:
        log({"step": "Dispatch"})
        jv = win32com.client.Dispatch("JVDTLab.JVLink")

        log({"step": "JVInit"})
        jv.JVInit(0)

        log({"step": "JVSetSavePath"})
        jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")

        log({"step": "JVSetSaveFlag"})
        jv.JVSetSaveFlag(1)

        log({"step": "JVSetPayFlag"})
        jv.JVSetPayFlag(0)

        log({"step": "JVOpen"})
        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)
        log({"step": "JVOpen_ret", "ret": int(open_ret), "readcount": int(readcount), "downloadcount": int(downloadcount), "lastts": str(lastts)})

        deadline = time.time() + 30.0
        i = 0
        while time.time() < deadline:
            i += 1
            st = int(jv.JVStatus())
            ret, buff, size, filename = invoke_read(jv)
            log({"step": "JVRead", "i": i, "status": st, "ret": int(ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:80]})
            if int(ret) == 0 and int(size) > 0:
                log({"step": "FOUND", "size": int(size), "filename": str(filename)})
                break
            time.sleep(0.5)

        log({"step": "JVClose"})
        jv.JVClose()

        log({"step": "done"})
    except Exception:
        log({"step": "exception", "traceback": traceback.format_exc()})
        raise


if __name__ == "__main__":
    main()