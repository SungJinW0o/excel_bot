# Changelog

All notable changes to this project will be documented in this file.

## [1.0.1] - 2026-02-14

- Rebuild desktop GUI layout with splitter + scrollable control panel + card sections to prevent text/field overlap.
- Improve readability and spacing across headings, labels, inputs, and action buttons.
- Add reliable no-console GUI launch path via `run_bot_gui.vbs` and strengthen `run_bot_gui.bat` diagnostics.
- Include new GUI launcher in installer payload and installer completion output.

## [1.0.0] - 2026-02-14

- Promote the project to a stable release.
- Include runtime hardening improvements from `0.1.5` and installer reliability fixes from `0.1.6`.
- Publish stable artifacts for installer and Python package distribution.

## [0.1.6] - 2026-02-14

- Fix Windows installer venv bootstrap by auto-repairing broken/missing pip before dependency installation.
- Retry pip install commands after pip repair to reduce setup failures on end-user machines.
- Add installer regression tests covering pip repair and retry flows.

## [0.1.5] - 2026-02-14

- Improve runtime stability by falling back to system Python when local `venv`/`.venv` is missing required runtime packages.
- Prevent duplicate row inflation on repeated runs by de-duplicating cleaned data before writing outputs.
- Change non-strict missing-SMTP behavior from `EMAIL_FAILED` to `EMAIL_SKIPPED` to reduce false-failure noise while keeping strict mode enforcement.
- Update setup workflow script to freeze environment to `requirements.lock.txt` instead of overwriting `requirements.txt`.

## [0.1.4] - 2026-02-13

- Fix GitHub release workflow PyPI publish gating and token handling.
- Bump package metadata version to 0.1.4.

## [0.1.0] - 2026-02-06

- Initial release.
- Data cleaning and report generation.
- Event logging for pipeline stages.
- DRY_RUN email notifications with attachment metadata.
- Cross-platform CLI and runner.
