from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def _try(call_desc: str, fn):
    try:
        r = fn()
        return {"ok": True, "call": call_desc, "result": str(r)}
    except Exception as e:
        return {"ok": False, "call": call_desc, "error": str(e)}


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    # JVInit は実装によりシグネチャが違うことがあるので候補を試す
    candidates = []

    # 0引数
    candidates.append(_try("JVInit()", lambda: jv.JVInit()))

    # 1引数（空/0/文字列）
    candidates.append(_try("JVInit(0)", lambda: jv.JVInit(0)))
    candidates.append(_try("JVInit('')", lambda: jv.JVInit("")))
    candidates.append(_try("JVInit('0')", lambda: jv.JVInit("0")))

    # 2引数（よくある：アプリ名・フラグ等）
    candidates.append(_try("JVInit('keiba_scraping', 0)", lambda: jv.JVInit("keiba_scraping", 0)))
    candidates.append(_try("JVInit('keiba_scraping', '')", lambda: jv.JVInit("keiba_scraping", "")))
    candidates.append(_try("JVInit('', 0)", lambda: jv.JVInit("", 0)))

    # 3引数（アプリ名/ユーザー/フラグ等のケース）
    candidates.append(_try("JVInit('keiba_scraping', 'keiba_scraping', 0)", lambda: jv.JVInit("keiba_scraping", "keiba_scraping", 0)))
    candidates.append(_try("JVInit('', '', 0)", lambda: jv.JVInit("", "", 0)))

    # 成功した最初のものを���う
    success = next((c for c in candidates if c["ok"]), None)

    payload = {
        "ok": success is not None,
        "success": success,
        "candidates": candidates,
        "prog_id": "JVDTLab.JVLink",
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())