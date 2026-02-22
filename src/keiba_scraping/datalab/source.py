from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from keiba_scraping.data.source import RaceCardSource


@dataclass(frozen=True)
class DataLabRaceCardSource(RaceCardSource):
    python32_path: str
    repo_root: Path

    def _run_32bit(self, rel_script_path: str) -> dict[str, Any]:
        script = (self.repo_root / rel_script_path).resolve()
        if not script.exists():
            raise FileNotFoundError(f"Missing 32-bit helper script: {script}")

        proc = subprocess.run(
            [self.python32_path, str(script)],
            capture_output=True,
            text=True,
            check=False,
        )
        out = (proc.stdout or "").strip()
        if not out:
            raise RuntimeError(f"32-bit helper returned empty stdout. stderr={proc.stderr!r}")

        try:
            payload = json.loads(out)
        except Exception as e:
            raise RuntimeError(f"Failed to parse helper JSON: {out!r}") from e

        payload["_returncode"] = proc.returncode
        if proc.stderr:
            payload["_stderr"] = proc.stderr.strip()
        return payload

    def get_race_card(self, race_id: str):
        # まずは疎通だけ。成功したらここから先を実装していく。
        payload = self._run_32bit("tools/jvlink32/jvlink_ping.py")
        if not payload.get("ok"):
            raise RuntimeError(f"JV-Link ping failed: {payload}")

        raise NotImplementedError(
            "JV-Link connection OK. Next step: implement race card fetch/convert to RaceCard."
        )