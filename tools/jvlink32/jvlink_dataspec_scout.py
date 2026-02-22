from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def invoke_open(jv, dataspec: str, fromdate: str, option: int):
    # dispid=7 (JVOpen)
    return jv._oleobj_.InvokeTypes(
        7, 0, 1,
        (3, 0),
        ((8, 1), (8, 1), (3, 1), (16387, 3), (16387, 3), (16392, 2)),
        dataspec, fromdate, option, 0, 0, ""
    )


def invoke_read(jv):
    # dispid=9 (JVRead)
    return jv._oleobj_.InvokeTypes(
        9, 0, 1,
        (3, 0),
        ((16392, 2), (16387, 2), (16392, 2)),
        "", 0, ""
    )


def main() -> int:
    fromdate = "20260222"
    option = 0

    # ありがちなデータ種別コード候補（まずは短いコード中心）
    # ※当たりを見つけるのが目的なので、ここは後で調整します
    candidates = [
        "RA", "RB", "RC", "RD", "RE", "RF",
        "SE", "SK", "SR", "SA",
        "YK", "YA",
        "KS", "KY",
        "TY", "TK",
    ]

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    jv.JVInit(0)
    jv.JVSetUIProperties()
    jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    jv.JVSetSaveFlag(1)

    results = []
    for ds in candidates:
        item = {"dataspec": ds}
        try:
            open_res = invoke_open(jv, ds, fromdate, option)
            item["open"] = str(open_res)
            read_res = invoke_read(jv)
            item["read"] = str(read_res)

            # readタプル文字列の中にデータが乗ってそうか（簡易）
            # read_res は (ret, buff, size, filename) の順で返ることが多い
            # ここでは文字列化してるので、buffが空でない雰囲気だ��見る
            results.append(item)
        except Exception as e:
            item["error"] = str(e)
            results.append(item)
        finally:
            try:
                jv.JVClose()
            except Exception:
                pass

    print(json.dumps({"ok": True, "fromdate": fromdate, "option": option, "results": results}, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())