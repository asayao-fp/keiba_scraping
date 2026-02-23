from __future__ import annotations

import gc
import json
import os
import sys
import time 

import pythoncom  # type: ignore
import win32com.client  # type: ignore

import pathlib

LOG_PATH = pathlib.Path(__file__).with_name("jvlink_open_debug.trace.jsonl")

def log(obj: dict) -> None:
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")
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

# ...（import と invoke_open/invoke_read はそのまま）...

def main() -> int:
    pythoncom.CoInitialize()
    jv = None

    close_ret = None
    payload: dict = {"ok": False}

    try:
        dataspec = os.environ.get("JV_DATASPEC", "RACE")
        fromdate = os.environ.get("JV_FROMDATE", "20240101000000")
        option = int(os.environ.get("JV_OPTION", "1"))

        # リトライ設定（環境変数で上書き可にすると便利）
        max_wait_sec = float(os.environ.get("JV_READ_MAX_WAIT_SEC", "60"))
        interval_sec = float(os.environ.get("JV_READ_INTERVAL_SEC", "0.5"))

        jv = win32com.client.Dispatch("JVDTLab.JVLink")
        log({"step": "dispatch_ok"})
        init_ret = jv.JVInit(0)

        ui_ret = None  # UIは呼ばない

        savepath_ret = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
        saveflag_ret = jv.JVSetSaveFlag(1)

        st0 = jv.JVStatus()

        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)
        log({"step": "open_ok", "open_ret": int(open_ret), "downloadcount": int(downloadcount)})

        st1 = jv.JVStatus()

        # ここから Read をリトライ
        attempts = []
        found = False
        read_ret = -9999
        buff = ""
        size = 0
        filename = ""

        deadline = time.time() + max_wait_sec

        log({"step": "read_try2", "status": st1, "ret": 1, "size": int(size)})

        try:
            while time.time() < deadline:
                st = int(jv.JVStatus())
                #read_ret, buff, size, filename = invoke_read(jv)
                read_ret,buff,size,filename = 1,"",0,""  # ダミー値（実際の呼び出しは上の行）
                attempts.append(
                    {
                        "status": st,
                        "ret": int(read_ret),
                        "size": int(size),
                        "filename": str(filename),
                        "buff_head": str(buff)[:30],
                    }
                )

                log({"step": "read_try", "status": st, "ret": int(read_ret), "size": int(size)})
                if int(read_ret) == 0 and int(size) > 0:
                    found = True
                    break

                # -3 は待機して再試行
                if int(read_ret) == -3:
                    time.sleep(interval_sec)
                    continue
                # それ以外の負数はエラー扱いで中断
                break
        except Exception as e:
            log({"step": "read_exception", "error": repr(e)})
            raise

        st2 = jv.JVStatus()

        payload = {
            "ok": True,
            "setup": {
                "init": int(init_ret),
                "ui": ui_ret,
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
                "found": bool(found),
                "ret": int(read_ret),
                "size": int(size),
                "filename": str(filename),
                "buff_head": str(buff)[:200],
                "attempts_tail": attempts[-10:],  # 直近10回だけ
            },
            "close": None,
        }
        return 0

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