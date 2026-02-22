from __future__ import annotations

import json
import sys

import win32com.client  # type: ignore


def _mask(s: str) -> str:
    if not s:
        return ""
    if len(s) <= 6:
        return "*" * len(s)
    return s[:3] + "*" * (len(s) - 5) + s[-2:]


def main() -> int:
    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    before = getattr(jv, "m_servicekey", "")
    init_ret = jv.JVInit(0)
    after_init = getattr(jv, "m_servicekey", "")

    payload = {
        "ok": True,
        "init_ret": str(init_ret),
        "m_servicekey_before": {"len": len(before), "masked": _mask(str(before))},
        "m_servicekey_after_init": {"len": len(after_init), "masked": _mask(str(after_init))},
        "m_savepath": getattr(jv, "m_savepath", None),
        "m_saveflag": getattr(jv, "m_saveflag", None),
    }
    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())