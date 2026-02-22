from __future__ import annotations

import json
import sys
import time

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
    dataspec = "RACE"
    fromdate = "20240101000000"
    option = 1

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.JVInit(0)
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)
    jv.JVSetPayFlag(0)

    open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)

    attempts = []
    record = None

    # 最大30秒待つ（-3の間は待機して再試行）
    deadline = time.time() + 30.0
    while time.time() < deadline:
        st = jv.JVStatus()
        ret, buff, size, filename = invoke_read(jv)
        attempts.append(
            {
                "status": int(st),
                "ret": int(ret),
                "size": int(size),
                "filename": str(filename),
                "buff_head": str(buff)[:80],
            }
        )

        if int(ret) == 0 and int(size) > 0:
            record = {"size": int(size), "filename": str(filename), "buff": str(buff)}
            break

        # -3 は「待て」系を想定、-203 なら未オープン/対象なし、その他負数も一旦待って様子見
        time.sleep(0.5)

    jv.JVClose()

    payload = {
        "ok": True,
        "open": {
            "ret": int(open_ret),
            "readcount": int(readcount),
            "downloadcount": int(downloadcount),
            "lastfiletimestamp": str(lastts),
        },
        "attempts_tail": attempts[-10:],  # 直近10回分だけ
        "record_found": record is not None,
        "record_preview": (record["buff"][:200] if record else ""),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())