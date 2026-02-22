from __future__ import annotations

import json
import sys
import time
import traceback

import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def invoke_read(jv):
    # これが落ちる可能性があるので、落ちたらプロセスごと終了しうる
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    try:
        dataspec = sys.argv[1] if len(sys.argv) > 1 else "RACE"
        fromdate = sys.argv[2] if len(sys.argv) > 2 else "20240101000000"
        option = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        sleep_sec = int(sys.argv[4]) if len(sys.argv) > 4 else 5

        jv = win32com.client.Dispatch("JVDTLab.JVLink")
        jv.JVInit(0)
        jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
        jv.JVSetSaveFlag(1)
        jv.JVSetPayFlag(0)

        open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

        time.sleep(sleep_sec)

        ret, buff, size, filename = invoke_read(jv)

        jv.JVClose()

        out = {
            "ok": True,
            "open": {"ret": int(open_ret), "readcount": int(readcount), "downloadcount": int(downloadcount), "lastts": str(lastts)},
            "read": {"ret": int(ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:120]},
        }
        print(json.dumps(out, ensure_ascii=False))
        return 0
    except Exception:
        print(json.dumps({"ok": False, "error": traceback.format_exc()}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())