from __future__ import annotations

import json
import sys
import time
import traceback

import win32com.client  # type: ignore


def log(step: str, extra=None):
    rec = {"t": time.strftime("%H:%M:%S"), "step": step}
    if extra is not None:
        rec["extra"] = extra
    print(json.dumps(rec, ensure_ascii=False), flush=True)


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


def run_one(payflag: int, dataspec: str, fromdate: str, option: int):
    log("Dispatch")
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    log("JVInit")
    init_ret = jv.JVInit(0)

    # 注意：ここでは UI を出さない（設定画面が出てハングしやすいので）
    # log("JVSetUIProperties")
    # ui_ret = jv.JVSetUIProperties()

    log("JVSetSavePath")
    sp_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")

    log("JVSetSaveFlag")
    sf_ret = jv.JVSetSaveFlag(1)

    log("JVSetPayFlag", payflag)
    pf_ret = jv.JVSetPayFlag(payflag)

    log("JVOpen", {"dataspec": dataspec, "fromdate": fromdate, "option": option})
    open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

    log("JVRead")
    read_ret, buff, size, filename = invoke_read(jv)

    log("JVClose")
    close_ret = jv.JVClose()

    return {
        "setup": {"init": int(init_ret), "savepath": int(sp_ret), "saveflag": int(sf_ret), "payflag_ret": int(pf_ret)},
        "open": {"ret": int(open_ret), "readcount": int(readcount), "downloadcount": int(downloadcount), "lastts": str(lastts)},
        "read": {"ret": int(read_ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:120]},
        "close": int(close_ret),
    }


def main() -> int:
    # まずはケースを1つに絞ってハングしないか確認
    payflag = 0
    dataspec = "RACE"
    option = 1
    fromdate = "20240101000000"

    try:
        result = run_one(payflag, dataspec, fromdate, option)
        log("DONE", result)
    except Exception:
        log("EXCEPTION", traceback.format_exc())
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())