from __future__ import annotations

import json
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


def run_case(payflag: int, dataspec: str, fromdate: str, option: int):
    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.ParentHWnd = win32gui.GetForegroundWindow()

    jv.JVInit(0)
    jv.JVSetUIProperties()
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)
    pay_ret = jv.JVSetPayFlag(payflag)

    open_ret, readcount, downloadcount, lastts = invoke_open(jv, dataspec, fromdate, option)
    read_ret, buff, size, filename = invoke_read(jv)

    jv.JVClose()

    return {
        "payflag": payflag,
        "pay_ret": int(pay_ret),
        "open": {
            "ret": int(open_ret),
            "readcount": int(readcount),
            "downloadcount": int(downloadcount),
            "lastfiletimestamp": str(lastts),
        },
        "read": {
            "ret": int(read_ret),
            "size": int(size),
            "filename": str(filename),
            "buff_head": str(buff)[:120],
        },
    }


def main() -> int:
    # 未来ではなく「過去日」をまず試す（データ存在の確率が高い日）
    # 日付は必要なら後で変えてOK
    cases = [
        {"dataspec": "RACE", "fromdate": "20250101000000", "option": 1},
        {"dataspec": "RACE", "fromdate": "20240101000000", "option": 1},
        {"dataspec": "RACE", "fromdate": "20240101000000", "option": 2},
    ]

    out = []
    for c in cases:
        for payflag in (0, 1):
            out.append(run_case(payflag, c["dataspec"], c["fromdate"], c["option"]))

    print(json.dumps({"ok": True, "results": out}, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())