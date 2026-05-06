# Guild Stats Scraper

Scrapes weekly guild stats from [rockymoon.com](https://www.rockymoon.com) and saves them to a local spreadsheet. Each run logs a dated snapshot and shows what changed since the last run.

## Setup

No manual setup needed. Just run the script for your platform and it handles everything.

**Linux / Mac**
```bash
bash run.sh
```

**Windows** — double-click `run.bat`, or from a terminal:
```
run.bat
```

On the first run it will create a virtual environment and install dependencies automatically. Subsequent runs just fetch and update.

## Output

A file called `guild_stats.xlsx` is created in the same folder. It has two sheets:

- **Changes** — always shows the latest snapshot compared to the previous run. Members are sorted by economy. Gains are green, losses are red. New members are marked `NEW`, departed members appear greyed out with `LEFT`.
- **Data** — an append-only log of every run. One row per member per run, timestamped. Useful for looking back at any point in history.

A summary is also printed to the terminal each run.

## Changing the guild

Open `scrape.py` and edit the two lines at the top:

```python
GUILD_URL = "https://www.rockymoon.com/bob/Charts/Borealis/Guild/Index/91"
OUTPUT_FILE = Path("guild_stats.xlsx")
```

Replace `91` with the target guild's index number from the URL. Change the filename if you want a separate spreadsheet per guild.

## How often to run

There's no scheduler — just run it whenever you want a new snapshot. Weekly works well. The longer the gap between runs, the more meaningful the deltas in the Changes sheet.

## Requirements

- Python 3.10+
- Internet access to reach rockymoon.com
