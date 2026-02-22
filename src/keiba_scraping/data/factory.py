from __future__ import annotations

from keiba_scraping.data.source import RaceCardSource
from keiba_scraping.data.stub_source import StubRaceCardSource


def create_source(source_name: str) -> RaceCardSource:
    source_name = source_name.lower().strip()
    if source_name == "stub":
        return StubRaceCardSource()

    if source_name == "datalab":
        raise NotImplementedError("DataLab source is not implemented yet.")

    raise ValueError(f"Unknown source: {source_name}")