from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class HorseEntry:
    horse_id: str
    name: str
    # 3着以内確率（0〜1）
    p_top3: float


@dataclass(frozen=True)
class RaceCard:
    race_id: str
    horses: list[HorseEntry]