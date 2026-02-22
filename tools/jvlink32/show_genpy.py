from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def main() -> int:
    jv = win32com.client.gencache.EnsureDispatch("JVDTLab.JVLink")
    # gen_py の Python ファイル実体（型定義が書かれている）
    gen_module = sys.modules[jv.__class__.__module__]
    path = getattr(gen_module, "__file__", None)

    payload = {
        "ok": True,
        "genpy_module": jv.__class__.__module__,
        "genpy_file": path,
        "class_name": jv.__class__.__name__,
        "dir_hint": "Open genpy_file and search for 'def JVOpen' and 'def JVRead'",
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())