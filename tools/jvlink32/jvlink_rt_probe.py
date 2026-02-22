from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def invoke_read(jv):
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.JVInit(0)
    jv.JVSetUIProperties()
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)

    # RTはデータ種別が別（例：速報系）で、keyも固定文字列の場合がある
    # ここでは key を空/0/日付 で試して「受理される形」を探す
    candidates = [
        {"dataspec": "RA", "key": ""},     # 仮（当たらなければ次で dataspec を変える）
        {"dataspec": "RA", "key": "0"},
        {"dataspec": "RA", "key": "20260222"},
    ]

    out = []
    for c in candidates:
        ds, key = c["dataspec"], c["key"]
        item = {"dataspec": ds, "key": key}
        try:
            rt_ret = jv.JVRTOpen(ds, key)
            item["rt_open_ret"] = str(rt_ret)
            read_ret, buff, size, filename = invoke_read(jv)
            item["read_ret"] = int(read_ret)
            item["size"] = int(size)
            item["filename"] = str(filename)
            item["buff_head"] = str(buff)[:200]
        except Exception as e:
            item["error"] = str(e)
        finally:
            try:
                jv.JVClose()
            except Exception:
                pass
        out.append(item)

    print(json.dumps({"ok": True, "results": out}, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())