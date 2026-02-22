from __future__ import annotations

from keiba_scraping.data.source import RaceCardSource
from keiba_scraping.domain.models import HorseEntry, RaceCard


class StubRaceCardSource(RaceCardSource):
    def get_race_card(self, race_id: str) -> RaceCard:
        # MVP用の仮データ（後でDataLab/JV-Linkに差し替え）
        horses = [
            HorseEntry("H01", "HorseA", 0.62),
            HorseEntry("H02", "HorseB", 0.55),
            HorseEntry("H03", "HorseC", 0.49),
            HorseEntry("H04", "HorseD", 0.44),
            HorseEntry("H05", "HorseE", 0.41),
            HorseEntry("H06", "HorseF", 0.36),
            HorseEntry("H07", "HorseG", 0.31),
            HorseEntry("H08", "HorseH", 0.28),
        ]
        return RaceCard(race_id=race_id, horses=horses)