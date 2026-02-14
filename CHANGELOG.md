# Changelog

All notable changes to this project will be documented in this file.

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
