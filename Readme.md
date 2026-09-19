# Global launch analysis, 2020 to 2026

An analysis of 1,500 completed rocket launches, using data from [Launch Library 2](https://thespacedevs.com/llapi) by The Space Devs. The window runs from part of 2020 to 19 September 2026. It covers the most recent launches in the database, not the full archive, so read the numbers as a snapshot of this period.

The full write-up is in [`reports/global_launch_activity_2002-2026.pdf`](reports/global_launch_activity_2020-2026.pdf).

## Findings

- Launch volume more than doubled from about 160 in 2021 to 341 in 2025.
- SpaceX flew 42% of the launches. Falcon 9 flew 600 and failed once.
- 95.6% of launches succeeded. The twelve rockets with 20 or more launches succeeded 99.2% of the time, against 88.4% for the rest.
- Low Earth orbit was the destination for more than half of all launches.

## Run it

Requires [uv](https://docs.astral.sh/uv/).

```
uv sync
uv run python fetch_data.py
```

The free API tier limits requests, so downloading everything takes several runs. The script saves each page it fetches and resumes where it stopped.

Then open `analysis.ipynb` and run all cells to clean the data, and build the Word report with:

```
uv run python make_report.py
```

## Project layout

```
fetch_data.py      downloads launches from the API and writes data/launches.csv
analysis.ipynb     cleans the data and writes data/launches_clean.csv
make_report.py     builds a Word report with charts from the cleaned data
data/              launch data (raw API pages are not tracked)
reports/           the finished report
```

## Data notes

A launch counts as a success only when its status is Launch Successful. Partial failures count as failures. The database includes suborbital flights, such as Blue Origin's New Shepard, so totals cover more than orbital launches.

Data: Launch Library 2, The Space Devs.