# keiba_scraping

## Setup (Windows venv)
```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
pip install -e .
```

## Run (MVP)
```powershell
python .\scripts\predict.py --race-id TEST_RACE --select 5
```

- `--select 5` の場合、3連複 **5頭BOX = 10点** を出力します
- DataLab/JV-Link 連携は次段で `keiba_scraping/datalab/` に実装して差し替えます