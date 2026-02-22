from __future__ import annotations

import itertools
from dataclasses import dataclass

from keiba_scraping.domain.models import HorseEntry


@dataclass(frozen=True)
class TrifectaCombo:
    horse_ids: tuple[str, str, str]
    horse_names: tuple[str, str, str]
    # ここでは簡易スコア（組合せの強さ）として p_top3 の積を採用
    score: float


def make_trifecta_box(horses: list[HorseEntry]) -> list[TrifectaCombo]:
    combos: list[TrifectaCombo] = []
    for a, b, c in itertools.combinations(horses, 3):
        combos.append(
            TrifectaCombo(
                horse_ids=(a.horse_id, b.horse_id, c.horse_id),
                horse_names=(a.name, b.name, c.name),
                score=a.p_top3 * b.p_top3 * c.p_top3,
            )
        )
    combos.sort(key=lambda x: x.score, reverse=True)
    return combos