"""Setup script for the excel_bot package."""

# pylint: disable=import-error
from pathlib import Path

from setuptools import find_packages, setup


ROOT = Path(__file__).parent
README = ROOT / "README.md"

LONG_DESCRIPTION = ""
if README.exists():
    LONG_DESCRIPTION = README.read_text(encoding="utf-8")


setup(
    name="excel_bot",
    version="0.1.0",
    description="Cross-platform Excel Automation Bot with logging, events, and notifications",
    long_description=LONG_DESCRIPTION,
    long_description_content_type="text/markdown",
    author="SungJinWoo",
    license="MIT",
    packages=find_packages(include=["excel_bot", "excel_bot.*"]),
    include_package_data=True,
    package_data={"excel_bot": ["config.json", "users.json"]},
    install_requires=[
        "pandas>=2.0",
        "openpyxl>=3.1",
    ],
    python_requires=">=3.9",
    entry_points={
        "console_scripts": [
            "excel-bot=excel_bot.run_bot:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: MIT License",
    ],
)
