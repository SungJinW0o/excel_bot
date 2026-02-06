# Excel Bot

[![MIT License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

Cross-platform Excel automation bot for cleaning, summarizing, and event logging.
Supports notifications, DRY_RUN mode, and optional email.

## Install

```bash
pip install dist/excel_bot-0.1.0-py3-none-any.whl
```

## Usage

Dry-run mode (default, safe test):

```bash
excel-bot --dry-run true
```

Real run (emails enabled if configured):

```bash
excel-bot --dry-run false
```

Optional launcher:

```bash
python run_bot.py --dry-run true
```

Headless mode (do not open report/log files):

```bash
excel-bot --dry-run true --no-open
# or
excel-bot --dry-run true --headless
```

## Run from source

You can run directly from the repo without building a wheel:

```bash
python run_bot.py --dry-run true
```

Outputs go to `output_data/` and logs go to `logs/events.jsonl` by default.
Use `--no-open` or `--headless` to avoid opening files in server or CI runs.

Headless run from root:

```bash
python run_bot.py --dry-run true --no-open --headless
```

## Folder structure

- input_data/ - place Excel files here
- output_data/ - cleaned data and reports
- config.json - bot settings (optional overrides via EXCEL_BOT_CONFIG)
- users.json - users/admins (optional overrides via EXCEL_BOT_USERS)

## Features

- Data cleaning and validation
- Summary reports (overall, category, region)
- Event logging for all pipeline stages
- DRY_RUN mode for safe testing
- Optional email notifications
- Cross-platform CLI

## Safety notes

- Use virtual environments for isolation
- DRY_RUN mode prevents actual emails
- All outputs go to output_data/ to avoid system folder writes

## Build and distribute

```bash
pip install --upgrade build setuptools wheel
python -m build
```

Note: building a wheel requires `setuptools` and `wheel` to be installed.

Install locally for testing:

```bash
pip install dist/excel_bot-0.1.0-py3-none-any.whl
```

Run the CLI:

```bash
excel-bot --dry-run true
```
