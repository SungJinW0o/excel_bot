# Excel Bot

[![MIT License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

Cross-platform Excel automation bot for cleaning, summarizing, and event logging.
Supports notifications, DRY_RUN mode, and optional email.

## Install

```bash
pip install dist/excel_bot-1.0.0-py3-none-any.whl
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

Optional one-command setup/test/package helper:

```bat
setup_test_package.bat
```

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

Windows Excel integration:

- When open-files mode is enabled, `summary_report.xlsx` is opened in Microsoft Excel when Excel is detected.
- If Excel is not detected, the report opens with your default file association.
- You can force a specific Excel binary path with `EXCEL_BOT_EXCEL_PATH` (for example `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`).

## Run from source

You can run directly from the repo without building a wheel:

```bash
python run_bot.py --dry-run true
```

Outputs go to `output_data/` and logs go to `logs/events.jsonl` by default.
Use `--no-open` or `--headless` to avoid opening files in server or CI runs.
Required input columns: `Quantity`, `UnitPrice`, `Status`, `Category`, `Region`.
Optional benchmark column: `Expense` (defaults to `0` when missing).
If rows/files are rejected during validation, details are saved to `output_data/data_quality_issues.csv`.
The CLI exits with code `2` when no valid input rows are available after validation.
Rerunning the same input is de-duplicated before writing `cleaned_master.xlsx` (uses `OrderID` when configured).

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
- Benchmark analysis saved to report sheets and `output_data/benchmark_summary.csv`
- Financial metrics: Total Earning, Expenses, Savings, Savings Rate
- Executive dashboard visuals in `summary_report.xlsx` (KPI cards, category/region charts, savings-rate trend)
- Data quality diagnostics sheet (`Data_Quality_Issues`) when files/rows are skipped
- Event logging for all pipeline stages
- DRY_RUN mode for safe testing
- Optional email notifications
- Rerun-safe output de-duplication to prevent repeated row inflation
- Cross-platform CLI

## Safety notes

- Use virtual environments for isolation
- DRY_RUN mode prevents actual emails
- Email notifications are fail-safe by default (pipeline continues even if SMTP is not configured and logs `EMAIL_SKIPPED`)
- Set `EXCEL_BOT_STRICT_EMAIL=true` to make SMTP/email failures fail the run
- All outputs go to output_data/ to avoid system folder writes

## Licensing and Commercial Use

- This project is licensed under MIT (`LICENSE`), which allows commercial use, resale, and donation-supported distribution.
- You must keep the MIT copyright/permission notice in distributed copies or substantial portions.
- You must also follow third-party dependency licenses.
- See `THIRD_PARTY_LICENSES.md` for dependency license notices and verification guidance.

## Build and distribute

```bash
pip install --upgrade build setuptools wheel
python -m build
```

Note: building a wheel requires `setuptools` and `wheel` to be installed.

Install locally for testing:

```bash
pip install dist/excel_bot-1.0.0-py3-none-any.whl
```

Run the CLI:

```bash
excel-bot --dry-run true
```

## Packaging and Publishing

GitHub does not provide a native Python package registry flow like npm/nuget menus.
For this project, use:

- GitHub Releases for downloadable artifacts (`.whl`, `.tar.gz`, `setup.exe`)
- PyPI for Python package installation via `pip`

Automated workflow:

- `.github/workflows/release.yml` builds and publishes packaging artifacts
- Trigger by pushing a version tag (example: `v1.0.0`)
- Optional PyPI publish runs from `workflow_dispatch` when `publish_to_pypi=true` (Trusted Publishing/OIDC)

Tag and publish example:

```bash
git tag v1.0.0
git push origin v1.0.0
```
