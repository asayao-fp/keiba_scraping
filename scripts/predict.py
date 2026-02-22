from __future__ import annotations

import argparse

from keiba_scraping.app.predict import run_prediction


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--race-id", required=True, help="Race identifier.")
    parser.add_argument("--select", type=int, default=5, help="Number of horses to box (default=5 -> 10 combos).")
    parser.add_argument("--out", default="outputs/predictions.csv", help="Output CSV path.")
    parser.add_argument("--source", default="stub", choices=["stub", "datalab"], help="Data source backend.")
    args = parser.parse_args()

    run_prediction(race_id=args.race_id, select=args.select, out_path=args.out, source=args.source)


if __name__ == "__main__":
    main()