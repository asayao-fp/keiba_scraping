from __future__ import annotations

from pathlib import Path

from keiba_scraping.data.source import RaceCardSource
from keiba_scraping.data.stub_source import StubRaceCardSource


def create_source(source_name: str) -> RaceCardSource:
    source_name = source_name.lower().strip()
    if source_name == "stub":
        return StubRaceCardSource()

    if source_name == "datalab":
        from keiba_scraping.datalab.source import DataLabRaceCardSource

        # あなたの環境で確認済みの 32bit Python
        python32 = r"C:\Users\takuma_asayao\AppData\Local\Programs\Python\Python313-32\python.exe"

        # リポジトリルート推定（src/keiba_scraping/data/factory.py から3つ上）
        repo_root = Path(__file__).resolve().parents[3]

        return DataLabRaceCardSource(python32_path=python32, repo_root=repo_root)

    raise ValueError(f"Unknown source: {source_name}")