from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def main() -> int:
    # gencache経由だと型情報を読み込んでメソッド一覧/ヘルプが取りやすい
    jv = win32com.client.gencache.EnsureDispatch("JVDTLab.JVLink")

    # Pythonから見える属性名（メソッド/プロパティ）を列挙
    names = sorted({n for n in dir(jv) if not n.startswith("_")})

    # 特に見たいメソッド
    focus = ["JVInit", "JVOpen", "JVRead", "JVClose", "JVStatus", "JVSetUIProperties"]

    focus_doc = {}
    for f in focus:
        try:
            attr = getattr(jv, f)
            focus_doc[f] = {
                "exists": True,
                "repr": repr(attr),
                # __doc__ に COM のシグネチャが出ることが多い
                "doc": getattr(attr, "__doc__", None),
            }
        except Exception as e:
            focus_doc[f] = {"exists": False, "error": str(e)}

    payload = {
        "ok": True,
        "prog_id": "JVDTLab.JVLink",
        "focus": focus_doc,
        "all_public_names_sample": names[:80],  # 多いので先頭だけ
        "all_public_names_count": len(names),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())