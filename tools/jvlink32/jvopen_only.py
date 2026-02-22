from __future__ import annotations

import gc
import json
import os
import sys

import pythoncom  # type: ignore
import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def main() -> int:
    pythoncom.CoInitialize()
    jv = None

    payload = {"ok": False}
    close_ret = None
    exit_code = 0

    try:
        dataspec = os.environ.get("JV_DATASPEC", "RACE")
        fromdate = os.environ.get("JV_FROMDATE", "20240101000000")
        option = int(os.environ.get("JV_OPTION", "1"))

        jv = win32com.client.Dispatch("JVDTLab.JVLink")
        init_ret = jv.JVInit(0)
        jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
        jv.JVSetSaveFlag(1)
        jv.JVSetPayFlag(0)

        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

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
            "close": None,
        }
        return 0

    except Exception as e:
        # 例外が出てもJSONを出す
        payload = {"ok": False, "error": repr(e), "close": None}
        exit_code = 2
        return exit_code

    finally:
        try:
            if jv is not None:
                close_ret = int(jv.JVClose())
        except Exception:
            close_ret = -9999

        # 明示的に参照を切る → GC → COM uninit の順にする
        jv = None
        gc.collect()
        pythoncom.CoUninitialize()

        if isinstance(payload, dict):
            payload["close"] = close_ret

        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)


if __name__ == "__main__":
    raise SystemExit(main())