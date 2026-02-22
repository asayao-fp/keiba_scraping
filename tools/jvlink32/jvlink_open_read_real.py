from __future__ import annotations

import json
import os
import sys

import pythoncom  # type: ignore
import win32com.client  # type: ignore


def byref_i4(v: int = 0):
    return win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, v)


def byref_bstr(s: str = ""):
    return win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, s)


def main() -> int:
    dataspec = os.environ.get("JV_DATASPEC", "RACE")
    fromdate = os.environ.get("JV_FROMDATE", "20260222")
    option = int(os.environ.get("JV_OPTION", "0"))

    jv = win32com.client.gencache.EnsureDispatch("JVDTLab.JVLink")

    # 初期設定
    jv.JVInit(0)
    jv.JVSetUIProperties()
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)

    # ServiceKey はここでは触らない（必要なら次段で環境変数から渡す）
    # jv.JVSetServiceKey(...)

    readcount = byref_i4(0)
    downloadcount = byref_i4(0)
    lastfiletimestamp = byref_bstr("")

    open_ret = jv.JVOpen(dataspec, fromdate, option, readcount, downloadcount, lastfiletimestamp)

    buff = byref_bstr("")
    size = byref_i4(0)
    filename = byref_bstr("")
    read_ret = jv.JVRead(buff, size, filename)

    jv.JVClose()

    payload = {
        "ok": True,
        "open": {
            "ret": str(open_ret),
            "dataspec": dataspec,
            "fromdate": fromdate,
            "option": option,
            "readcount": int(readcount.value),
            "downloadcount": int(downloadcount.value),
            "lastfiletimestamp": str(lastfiletimestamp.value),
        },
        "read": {
            "ret": str(read_ret),
            "size": int(size.value),
            "filename": str(filename.value),
            # データが巨大な場合に備えて先頭だけ
            "buff_head": (str(buff.value)[:200] if buff.value is not None else ""),
        },
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())