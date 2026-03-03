# WKORPUSE Runbook

**Weekdaily KOR Purchase of US Equities** - Automated data collection from KSD (Korea Securities Depository) SEIBro website.

## Overview

This runbook automates the daily extraction of Korean investors' US equity purchase/settlement data from the [SEIBro portal](https://seibro.or.kr). It downloads the top 50 stocks by net purchase settlement amount, parses the data, and outputs clean XLS/CSV files.

## Data Source

- **Provider**: KSD - Korea Securities Depository
- **Portal**: SEIBro (Securities Information BROker)
- **URL**: `https://seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/ovsSec/BIP_CNTS10013V.xml&menuNo=921`
- **Frequency**: Weekdaily (Monday-Friday)
- **Country Filter**: USA
- **Currency**: USD

## Pipeline

```
orchestrator.py
  |
  +-- Step 1: scraper.py     - Download data from SEIBro (Selenium)
  +-- Step 2: parser.py      - Parse and clean HTML/Excel data
  +-- Step 3: file_generator.py - Generate output XLS/CSV files
```

## Date Logic

The system automatically determines the correct data date:

| Run Day    | Data Date |
|------------|-----------|
| Monday     | Friday    |
| Tuesday    | Monday    |
| Wednesday  | Tuesday   |
| Thursday   | Wednesday |
| Friday     | Thursday  |
| Saturday   | Friday    |
| Sunday     | Friday    |

Weekend data is not published. Override with `TARGET_DATE` in `config.py`.

## Setup

### Prerequisites

- Python 3.11+
- Google Chrome browser
- ChromeDriver (managed automatically by Selenium)

### Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Run the full pipeline

```bash
python orchestrator.py
```

### Run individual components

```bash
# Test scraper only
python scraper.py

# Test parser with a specific file
python parser.py path/to/file.xls
```

## Output

Files are saved to timestamped directories and a `latest/` folder:

```
output/
  20260202_163152/
    WKORPUSE_DATA_20260130.xls
    WKORPUSE_DATA_20260130.csv
  latest/
    WKORPUSE_DATA_latest.xls
    WKORPUSE_DATA_latest.csv
```

### Output Columns

| Korean     | English           | Type    |
|------------|-------------------|---------|
| 순위       | Rank              | integer |
| 국가       | Country           | string  |
| 종목코드   | Stock Code        | string  |
| 종목명     | Stock Name        | string  |
| 매수결제   | Buy Settlement    | numeric |
| 매도결제   | Sell Settlement   | numeric |
| 순매수결제 | Net Purchase      | numeric |

## Configuration

Key settings in `config.py`:

| Setting              | Default | Description                          |
|----------------------|---------|--------------------------------------|
| `TARGET_DATE`        | `None`  | Auto-calculate, or set a specific date |
| `HEADLESS_MODE`      | `False` | Run Chrome without UI                |
| `WAIT_TIMEOUT`       | `30`    | Selenium wait timeout (seconds)      |
| `PAGE_LOAD_DELAY`    | `5`     | Wait for dynamic content (seconds)   |
| `DOWNLOAD_WAIT_TIME` | `30`    | Excel download timeout (seconds)     |

## Project Structure

```
WKORPUSE_Runbook/
  orchestrator.py      # Main pipeline orchestrator
  scraper.py           # SEIBro web scraper (Selenium)
  parser.py            # Data parser (HTML/Excel)
  file_generator.py    # Output file generator
  config.py            # Configuration and selectors
  logger_setup.py      # Logging configuration
  requirements.txt     # Python dependencies
  project_information/ # Reference docs and samples
```
