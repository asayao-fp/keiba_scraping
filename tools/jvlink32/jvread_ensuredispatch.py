from __future__ import annotations

import json
import sys
import time

import win32com.client  # type: ignore


def main() -> int:
    jv = win32com.client.gencache.EnsureDispatch("JVDTLab.JVLink")
    jv.JVInit(0)
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)
    jv.JVSetPayFlag(0)

    # ここは InvokeTypes で open してもよいが、まずはラッパのJVOpenを使う
    # makepy定義を見る限り JVOpen(dataspec, fromdate, option) は (ret, readcount, downloadcount, lastts) を返す想定
    open_ret, readcount, downloadcount, lastts = jv.JVOpen("RACE", "20240101000000", 1)

    # 少し待ってからRead
    time.sleep(3)

    ret, buff, size, filename = jv.JVRead()

    jv.JVClose()

    payload = {
        "ok": True,
        "open": {"ret": int(open_ret), "readcount": int(readcount), "downloadcount": int(downloadcount), "lastts": str(lastts)},
        "read": {"ret": int(ret), "size": int(size), "filename": str(filename), "buff_head": str(buff)[:120]},
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())