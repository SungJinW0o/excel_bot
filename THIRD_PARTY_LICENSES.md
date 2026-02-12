# Third-Party License Notices

This project is licensed under the MIT License (see `LICENSE`).

When distributing this project (source, wheel, installer, or bundled app), you must also comply with licenses of third-party dependencies used by your distribution.

Core runtime dependencies used by this project:

- `pandas` - BSD 3-Clause
- `openpyxl` - MIT
- `PySide6` (optional GUI) - LGPLv3 / GPL-compatible Qt terms (or commercial Qt license, depending on your usage)

Notes:

- The exact dependency set can vary by install profile and version.
- Verify final installed package licenses before commercial distribution.

Quick check commands:

```bash
python -m pip show pandas openpyxl PySide6
```

For source-of-truth terms, refer to each dependency's official license files and project pages.
