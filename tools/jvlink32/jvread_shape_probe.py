from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def invoke_read_raw(jv):
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.JVInit(0)
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)
    jv.JVSetPayFlag(0)

    open_res = invoke_open(jv, "RACE", "20240101000000", 1)

    read_raw = invoke_read_raw(jv)

    jv.JVClose()

    payload = {
        "ok": True,
        "open_raw": repr(open_res),
        "read_raw_repr": repr(read_raw),
        "read_raw_type": str(type(read_raw)),
        "read_raw_len": (len(read_raw) if hasattr(read_raw, "__len__") else None),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())