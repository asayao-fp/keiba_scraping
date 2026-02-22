from __future__ import annotations

import csv
import os

from keiba_scraping.data.factory import create_source
from keiba_scraping.logic.trifecta_box import make_trifecta_box


def run_prediction(race_id: str, select: int, out_path: str, source: str = "stub") -> None:
    if select < 3:
        raise ValueError("--select must be >= 3")
    if select != 5:
        raise ValueError("MVP currently supports --select 5 only (10 tickets).")

    race_source = create_source(source)
    race = race_source.get_race_card(race_id)

    top = sorted(race.horses, key=lambda h: h.p_top3, reverse=True)[:select]

    combos = make_trifecta_box(top)
    if len(combos) != 10:
        raise RuntimeError(f"Expected 10 combos, got {len(combos)}")

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["race_id", "horse1", "horse2", "horse3", "score"])
        for c in combos:
            w.writerow([race_id, *c.horse_names, f"{c.score:.6f}"])

    print(f"race_id={race_id}")
    print(f"source={source}")
    print("selected horses:")
    for h in top:
        print(f"- {h.name} (p_top3={h.p_top3:.2f})")

    print("\n3連複 5頭BOX (10点):")
    for i, c in enumerate(combos, start=1):
        print(f"{i:02d}. {' - '.join(c.horse_names)}  score={c.score:.6f}")

    print(f"\nSaved: {out_path}")