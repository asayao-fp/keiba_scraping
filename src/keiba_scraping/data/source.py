from __future__ import annotations

from abc import ABC, abstractmethod

from keiba_scraping.domain.models import RaceCard


class RaceCardSource(ABC):
    @abstractmethod
    def get_race_card(self, race_id: str) -> RaceCard:
        raise NotImplementedError