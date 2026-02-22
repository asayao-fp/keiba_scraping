from __future__ import annotations

import json
import os
import re
import sys

import win32com.client  # type: ignore


def normalize(s: str) -> dict[str, str]:
    raw = s
    stripped = raw.strip()

    # 全角っぽいハイフンを半角に寄せる（よく混入するものだけ）
    stripped = stripped.replace("－", "-").replace("―", "-").replace("ー", "-")

    upper = stripped.upper()

    no_hyphen = re.sub(r"[^A-Z0-9]", "", upper)  # ハイフン含む記号全部除去して英数字だけ

    return {
        "raw_len": str(len(raw)),
        "stripped": stripped,
        "upper": upper,
        "no_hyphen": no_hyphen,
    }


def _try(jv, desc: str, key: str):
    try:
        r = jv.JVSetServiceKey(key)
        return {"ok": True, "call": desc, "ret": str(r), "key_len": len(key)}
    except Exception as e:
        return {"ok": False, "call": desc, "error": str(e), "key_len": len(key)}


def main() -> int:
    key = os.environ.get("JRAVAN_SERVICE_KEY", "")
    if not key:
        print(json.dumps({"ok": False, "error": "JRAVAN_SERVICE_KEY is not set"}, ensure_ascii=False), file=sys.stdout)
        return 2

    forms = normalize(key)

    jv = win32com.client.Dispatch("JVDTLab.JVLink")
    init = jv.JVInit(0)
    ui = jv.JVSetUIProperties()
    savepath = jv.JVSetSavePath(r"C:\ProgramData\JRA-VAN\Data")
    saveflag = jv.JVSetSaveFlag(1)

    results = [
        _try(jv, "JVSetServiceKey(stripped)", forms["stripped"]),
        _try(jv, "JVSetServiceKey(upper)", forms["upper"]),
        _try(jv, "JVSetServiceKey(no_hyphen)", forms["no_hyphen"]),
    ]

    payload = {
        "ok": True,
        "setup": {"init": str(init), "ui": str(ui), "savepath": str(savepath), "saveflag": str(saveflag)},
        "key_meta": {
            "raw_len": forms["raw_len"],
            "stripped_len": len(forms["stripped"]),
            "upper_len": len(forms["upper"]),
            "no_hyphen_len": len(forms["no_hyphen"]),
            # セキュリティのためキー自体は出さない
            "preview": forms["upper"][:4] + "..." + forms["upper"][-2:],
        },
        "results": results,
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())