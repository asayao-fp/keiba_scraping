"""jvread_via_bridge.py – Python wrapper that invokes JVLinkBridge.exe and
prints the returned JSON.

JVRead crashes the Python/pywin32 process with exit code 0xC0000409.
This wrapper delegates COM/JVRead work to the .NET CLI bridge, which runs in
a separate CLR process and is immune to the crash.

Usage
-----
# Run with defaults (reads RACE data from 2024-01-01):
python tools/jvlink32/jvread_via_bridge.py

# Override via environment variables:
set JV_DATASPEC=RACE
set JV_FROMDATE=20240101000000
set JV_OPTION=1
set JV_SAVE_PATH=C:\\ProgramData\\JRA-VAN\\Data
set JV_READ_MAX_WAIT_SEC=60
set JV_READ_INTERVAL_SEC=0.5

# Diagnostics / robustness env vars:
set JV_SLEEP_AFTER_OPEN_SEC=1.0
set JV_ENABLE_UI_PROPERTIES=1
set JV_ENABLE_STATUS_POLL=1
set JV_STATUS_POLL_MAX_WAIT_SEC=10
set JV_STATUS_POLL_INTERVAL_SEC=0.5
python tools/jvlink32/jvread_via_bridge.py

# Or pass positional args (dataspec fromdate option):
python tools/jvlink32/jvread_via_bridge.py RACE 20240101000000 1

Build the bridge first (PowerShell, x86):
  cd tools/jvlink32/JVLinkBridge
  dotnet build -c Release
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Locate the bridge executable (Debug or Release, x86)
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent
_BRIDGE_DIR = _HERE / "JVLinkBridge"

_CANDIDATES = [
    _BRIDGE_DIR / "bin" / "Release" / "net8.0-windows" / "win-x86" / "JVLinkBridge.exe",
    _BRIDGE_DIR / "bin" / "Release" / "net8.0-windows" / "JVLinkBridge.exe",
    _BRIDGE_DIR / "bin" / "Debug"   / "net8.0-windows" / "win-x86" / "JVLinkBridge.exe",
    _BRIDGE_DIR / "bin" / "Debug"   / "net8.0-windows" / "JVLinkBridge.exe",
]


def _find_bridge() -> Path:
    for p in _CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError(
        "JVLinkBridge.exe not found. Build it first:\n"
        "  cd tools/jvlink32/JVLinkBridge\n"
        "  dotnet build -c Release\n"
        f"Searched: {[str(c) for c in _CANDIDATES]}"
    )


def run_bridge(
    dataspec: str | None = None,
    fromdate: str | None = None,
    option: str | None = None,
    extra_env: dict[str, str] | None = None,
) -> dict:
    """Run JVLinkBridge and return the parsed JSON result dict."""
    bridge_exe = _find_bridge()

    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)

    args: list[str] = [str(bridge_exe)]
    if dataspec:
        args.append(dataspec)
    if fromdate:
        args.append(fromdate)
    if option:
        args.append(option)

    proc = subprocess.run(
        args,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )

    stderr_text = proc.stderr.strip()

    if proc.returncode not in (0, 1):
        raise RuntimeError(
            f"JVLinkBridge exited with unexpected code {proc.returncode}.\n"
            f"stderr: {stderr_text or '(empty)'}"
        )

    stdout = proc.stdout.strip()
    if not stdout:
        raise RuntimeError(
            f"JVLinkBridge produced no output (exit {proc.returncode}).\n"
            f"stderr: {stderr_text or '(empty)'}"
        )

    try:
        result = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"JVLinkBridge output is not valid JSON: {exc}\n"
            f"raw stdout: {stdout!r}\n"
            f"stderr: {stderr_text or '(empty)'}"
        ) from exc

    # Attach stderr to the result dict so callers can inspect it
    if stderr_text:
        result.setdefault("stderr", stderr_text)

    return result


def main() -> int:
    # Positional CLI args override env vars
    dataspec = sys.argv[1] if len(sys.argv) > 1 else None
    fromdate = sys.argv[2] if len(sys.argv) > 2 else None
    option   = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        result = run_bridge(dataspec=dataspec, fromdate=fromdate, option=option)
    except FileNotFoundError as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 2
    except RuntimeError as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 2

    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result.get("ok"):
        # Print a brief diagnostic hint to stderr so it doesn't pollute JSON stdout
        stage   = result.get("stage", "unknown")
        hresult = result.get("hresult", "")
        error   = result.get("error", "")
        hint = f"[jvread_via_bridge] Bridge failed at stage={stage!r}"
        if hresult:
            hint += f" hresult={hresult}"
        if error:
            hint += f": {error}"
        print(hint, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
