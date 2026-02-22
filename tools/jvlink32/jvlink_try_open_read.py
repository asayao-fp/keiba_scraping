from __future__ import annotations

import argparse
import json
import sys
from datetime import date, timedelta

import win32com.client  # type: ignore


def _today_yyyymmdd() -> str:
    return date.today().strftime("%Y%m%d")


def _read_some(jv, n: int = 3) -> list[str]:
    rows: list[str] = []
    for _ in range(n):
        r = jv.JVRead()
        rows.append(str(r))
    return rows


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="from_date", default=None, help="YYYYMMDD")
    ap.add_argument("--to", dest="to_date", default=None, help="YYYYMMDD")
    ap.add_argument("--n", dest="n", type=int, default=3, help="read rows")
    args = ap.parse_args()

    today = date.today()
    from_date = args.from_date or (today - timedelta(days=7)).strftime("%Y%m%d")
    to_date = args.to_date or (today + timedelta(days=21)).strftime("%Y%m%d")

    dataspec_candidates = [
        # まずは番組/レース情報っぽい候補
        "RACE", "RACEINFO", "RACE_INFO", "RACE_SCHEDULE", "SCHEDULE",
        "PROGRAM", "RACE_PROGRAM", "TOKUBETSU", "GRADED",
        # 出馬表系っぽい候補（番組が見つからない場合の保険）
        "CARD", "RACECARD", "RACE_CARD", "SYUSSO", "SHUSSO",
    ]

    jv = win32com.client.Dispatch("JVDTLab.JVLink")

    payload = {
        "ok": True,
        "from": from_date,
        "to": to_date,
        "today": _today_yyyymmdd(),
        "tried": [],
    }

    try:
        init_r = jv.JVInit(0)
        payload["jvinit"] = str(init_r)
    except Exception as e:
        payload["ok"] = False
        payload["error"] = f"JVInit(0) failed: {e}"
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
        return 2

    for ds in dataspec_candidates:
        item = {"dataspec": ds}
        try:
            try:
                r = jv.JVOpen(ds, from_date)
                item["open_call"] = "JVOpen(ds, from_date)"
                item["open_result"] = str(r)
            except Exception as e1:
                try:
                    r = jv.JVOpen(ds, from_date, to_date)
                    item["open_call"] = "JVOpen(ds, from_date, to_date)"
                    item["open_result"] = str(r)
                except Exception as e2:
                    item["ok"] = False
                    item["error"] = f"JVOpen failed (2-arg): {e1} / (3-arg): {e2}"
                    payload["tried"].append(item)
                    continue

            try:
                rows = _read_some(jv, n=args.n)
                item["ok"] = True
                item["read_rows"] = rows
                payload["tried"].append(item)
                break
            except Exception as e:
                item["ok"] = False
                item["error"] = f"JVRead failed: {e}"
                payload["tried"].append(item)
                continue
            finally:
                try:
                    jv.JVClose()
                except Exception:
                    pass

        except Exception as e:
            item["ok"] = False
            item["error"] = f"unexpected: {e}"
            payload["tried"].append(item)
            continue

    print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())