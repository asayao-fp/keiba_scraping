# keiba_scraping

## Setup (Windows venv)

cmd.exe example:

py -3.11 -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install -U pip
pip install -e .

## Run (MVP)

python .\scripts\predict.py --race-id TEST_RACE --select 5 --source stub

- --select 5 outputs 3連複 5頭BOX (10点)
- DataLab/JV-Link integration will be added under src/keiba_scraping/datalab
