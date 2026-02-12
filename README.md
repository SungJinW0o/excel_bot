# Excel Bot

[![MIT License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

Cross-platform Excel automation bot for cleaning, summarizing, and event logging.
Supports notifications, DRY_RUN mode, and optional email.

## Install

```bash
pip install dist/excel_bot-0.1.0-py3-none-any.whl
```

## Quick Start (Windows)

1. Put your Excel files in `input_data/`.
2. Run the simple launcher:

```bat
run_bot.bat
```

3. Choose mode:
- `1` = safe test run (`DRY_RUN=true`, recommended)
- `2` = live run (`DRY_RUN=false`)

## Desktop GUI (Glass UI)

Install GUI dependency:

```bash
pip install "excel_bot[gui]"
```

Launch GUI:

```bash
python run_bot_gui.py
# or
excel-bot-gui
```

Windows shortcut launcher:

```bat
run_bot_gui.bat
```

GUI options include:
- Safe test vs live mode
- Open/close report files after run
- Working directory selection
- Optional custom `config.json` and `users.json` selection
- Live run console and status badge
- Integrity verification (directories, config/users validation, dependency checks)
- One-click sample data generation
- App update option (pip upgrade source)
- Optional feature download (GUI/installer/test tooling)

## Usage (CLI)

Dry-run mode (default, safe test):

```bash
python run_bot.py --dry-run true
```

Real run (emails enabled if configured):

```bash
python run_bot.py --dry-run false
```

Headless mode (do not open report/log files):

```bash
python run_bot.py --dry-run true --no-open
# or
python run_bot.py --dry-run true --headless
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
