#!/usr/bin/env python3
"""Desktop UI for Excel Bot."""

from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .auth import load_users
from .config import load_config


try:
    from PySide6.QtCore import QEasingCurve, QObject, QPropertyAnimation, QThread, Qt, QUrl, Signal, Slot
    from PySide6.QtGui import QDesktopServices, QFont
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFrame,
        QGraphicsBlurEffect,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover - optional GUI dependency
    _GUI_IMPORT_ERROR = exc
else:
    _GUI_IMPORT_ERROR = None


STYLE_SHEET = """
QWidget {
    color: #e6edf8;
    font-family: "Segoe UI", "Helvetica Neue", sans-serif;
    font-size: 13px;
}
QMainWindow {
    background-color: #0a1020;
}
QFrame#Surface {
    background: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.20);
    border-radius: 18px;
}
QLabel#Title {
    font-size: 30px;
    font-weight: 700;
}
QLabel#Subtitle {
    color: rgba(230, 237, 248, 0.78);
    font-size: 14px;
}
QLabel#SectionTitle {
    font-size: 16px;
    font-weight: 600;
}
QLabel#StatusBadge {
    border-radius: 10px;
    padding: 4px 10px;
    font-weight: 600;
}
QLineEdit, QComboBox {
    background: rgba(12, 20, 36, 0.75);
    border: 1px solid rgba(255, 255, 255, 0.22);
    border-radius: 10px;
    padding: 8px 10px;
}
QLineEdit:focus, QComboBox:focus {
    border-color: rgba(122, 184, 255, 0.85);
}
QPushButton {
    border-radius: 12px;
    padding: 9px 12px;
    border: 1px solid rgba(255, 255, 255, 0.25);
    background: rgba(255, 255, 255, 0.12);
}
QPushButton:hover {
    background: rgba(255, 255, 255, 0.20);
}
QPushButton:disabled {
    background: rgba(255, 255, 255, 0.08);
    color: rgba(230, 237, 248, 0.45);
}
QPushButton#PrimaryButton {
    border: 0;
    background: qlineargradient(
        x1: 0, y1: 0, x2: 1, y2: 1,
        stop: 0 #27a7ff,
        stop: 1 #3e7bff
    );
    font-weight: 700;
}
QPushButton#PrimaryButton:hover {
    background: qlineargradient(
        x1: 0, y1: 0, x2: 1, y2: 1,
        stop: 0 #41b3ff,
        stop: 1 #4f88ff
    );
}
QPlainTextEdit {
    background: rgba(8, 14, 26, 0.75);
    border: 1px solid rgba(255, 255, 255, 0.16);
    border-radius: 12px;
    padding: 10px;
}
QProgressBar {
    border: 1px solid rgba(255, 255, 255, 0.2);
    border-radius: 8px;
    text-align: center;
    background: rgba(8, 14, 26, 0.85);
}
QProgressBar::chunk {
    border-radius: 7px;
    background: qlineargradient(
        x1: 0, y1: 0, x2: 1, y2: 0,
        stop: 0 #35c0ff,
        stop: 1 #6f7bff
    );
}
"""


FEATURE_PACKAGES: Dict[str, str] = {
    "GUI Components (PySide6)": "PySide6>=6.7",
    "Installer Builder (PyInstaller)": "pyinstaller>=6.18",
    "Testing Tools (pytest)": "pytest>=8.0",
}
DEFAULT_UPDATE_SOURCE = "git+https://github.com/SungJinW0o/excel_bot.git"


if _GUI_IMPORT_ERROR is None:
    class CommandWorker(QObject):
        """Run subprocess commands off the UI thread."""

        line_emitted = Signal(str)
        finished = Signal(int, str)

        def __init__(
            self,
            command: List[str],
            work_dir: Path,
            task_name: str,
            env: Optional[Dict[str, str]] = None,
        ) -> None:
            super().__init__()
            self.command = command
            self.work_dir = work_dir
            self.task_name = task_name
            self.env = env

        @Slot()
        def run(self) -> None:
            merged_env = os.environ.copy()
            if self.env:
                merged_env.update(self.env)

            self.line_emitted.emit(f"[{self.task_name}] Starting...")
            self.line_emitted.emit(f"[{self.task_name}] Command: {' '.join(self.command)}")
            self.line_emitted.emit(f"[{self.task_name}] Working directory: {self.work_dir}")
            self.line_emitted.emit("")

            process = subprocess.Popen(
                self.command,
                cwd=str(self.work_dir),
                env=merged_env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )

            assert process.stdout is not None
            for line in process.stdout:
                self.line_emitted.emit(line.rstrip("\n"))

            self.finished.emit(process.wait(), self.task_name)


    class BotWindow(QMainWindow):
        """Main GUI window."""

        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Excel Bot Studio")
            self.resize(1140, 740)
            self.setMinimumSize(1020, 660)
            self.setStyleSheet(STYLE_SHEET)
            self.setWindowOpacity(0.0)

            self._thread: Optional[QThread] = None
            self._worker: Optional[CommandWorker] = None

            self._root = Path.cwd()
            self._build_ui()
            self._set_status("idle", "Ready")
            self._play_intro_animation()

        def _build_ui(self) -> None:
            root = QWidget(self)
            self.setCentralWidget(root)

            self._orb_a = QLabel(root)
            self._orb_b = QLabel(root)
            self._orb_c = QLabel(root)
            self._setup_orb(self._orb_a, 280, "rgba(62, 153, 255, 185)")
            self._setup_orb(self._orb_b, 360, "rgba(116, 116, 255, 150)")
            self._setup_orb(self._orb_c, 240, "rgba(57, 212, 181, 130)")

            outer = QVBoxLayout(root)
            outer.setContentsMargins(24, 22, 24, 22)
            outer.setSpacing(16)

            title = QLabel("Excel Bot Studio")
            title.setObjectName("Title")
            subtitle = QLabel("Smooth pipeline control with integrity checks, updates, and feature installs.")
            subtitle.setObjectName("Subtitle")
            outer.addWidget(title)
            outer.addWidget(subtitle)

            content = QHBoxLayout()
            content.setSpacing(16)
            outer.addLayout(content, 1)

            left_panel = QFrame()
            left_panel.setObjectName("Surface")
            left_panel.setMinimumWidth(420)
            left_layout = QVBoxLayout(left_panel)
            left_layout.setContentsMargins(18, 18, 18, 18)
            left_layout.setSpacing(12)
            content.addWidget(left_panel, 0)

            right_panel = QFrame()
            right_panel.setObjectName("Surface")
            right_layout = QVBoxLayout(right_panel)
            right_layout.setContentsMargins(18, 18, 18, 18)
            right_layout.setSpacing(12)
            content.addWidget(right_panel, 1)

            left_layout.addWidget(self._section("Run Options"))

            self.work_dir_edit = QLineEdit(str(self._root))
            left_layout.addWidget(QLabel("Working directory"))
            left_layout.addLayout(self._path_row(self.work_dir_edit, self._browse_work_dir))

            left_layout.addWidget(QLabel("Configuration file (optional)"))
            self.config_edit = QLineEdit("")
            self.config_edit.setPlaceholderText("Use default config.json")
            left_layout.addLayout(self._path_row(self.config_edit, self._browse_config_file))

            left_layout.addWidget(QLabel("Users file (optional)"))
            self.users_edit = QLineEdit("")
            self.users_edit.setPlaceholderText("Use default users.json")
            left_layout.addLayout(self._path_row(self.users_edit, self._browse_users_file))

            left_layout.addWidget(QLabel("Run mode"))
            self.mode_combo = QComboBox()
            self.mode_combo.addItem("Safe test (DRY_RUN=true)")
            self.mode_combo.addItem("Live run (DRY_RUN=false)")
            left_layout.addWidget(self.mode_combo)

            self.open_files_check = QCheckBox("Open report/log files after run")
            self.open_files_check.setChecked(True)
            left_layout.addWidget(self.open_files_check)

            self.run_button = QPushButton("Run Excel Bot")
            self.run_button.setObjectName("PrimaryButton")
            self.run_button.clicked.connect(self._start_run)
            left_layout.addWidget(self.run_button)

            quick_actions = QHBoxLayout()
            input_btn = QPushButton("Open Input")
            output_btn = QPushButton("Open Output")
            logs_btn = QPushButton("Open Logs")
            self.sample_button = QPushButton("Generate Sample")
            input_btn.clicked.connect(lambda: self._open_relative_folder("input_data"))
            output_btn.clicked.connect(lambda: self._open_relative_folder("output_data"))
            logs_btn.clicked.connect(lambda: self._open_relative_folder("logs"))
            self.sample_button.clicked.connect(self._generate_sample_data)
            quick_actions.addWidget(input_btn)
            quick_actions.addWidget(output_btn)
            quick_actions.addWidget(logs_btn)
            quick_actions.addWidget(self.sample_button)
            left_layout.addLayout(quick_actions)

            left_layout.addWidget(self._section("Maintenance"))

            self.integrity_button = QPushButton("Verify Integrity")
            self.integrity_button.clicked.connect(self._verify_integrity)
            left_layout.addWidget(self.integrity_button)

            left_layout.addWidget(QLabel("Update source"))
            self.update_source_edit = QLineEdit(DEFAULT_UPDATE_SOURCE)
            left_layout.addWidget(self.update_source_edit)

            self.update_button = QPushButton("Update App")
            self.update_button.clicked.connect(self._update_app)
            left_layout.addWidget(self.update_button)

            left_layout.addWidget(QLabel("Download optional feature"))
            self.feature_combo = QComboBox()
            for feature_name in FEATURE_PACKAGES:
                self.feature_combo.addItem(feature_name)
            left_layout.addWidget(self.feature_combo)

            self.install_feature_button = QPushButton("Download Feature")
            self.install_feature_button.clicked.connect(self._install_feature)
            left_layout.addWidget(self.install_feature_button)
            left_layout.addStretch(1)

            header_row = QHBoxLayout()
            header_row.addWidget(self._section("Run Console"))
            header_row.addStretch(1)
            self.status_label = QLabel()
            self.status_label.setObjectName("StatusBadge")
            header_row.addWidget(self.status_label)
            right_layout.addLayout(header_row)

            self.progress = QProgressBar()
            self.progress.setVisible(False)
            right_layout.addWidget(self.progress)

            self.log_view = QPlainTextEdit()
            self.log_view.setReadOnly(True)
            self.log_view.setLineWrapMode(QPlainTextEdit.NoWrap)
            self.log_view.setFont(QFont("Cascadia Mono", 10))
            right_layout.addWidget(self.log_view, 1)

            self.summary_label = QLabel("Summary: waiting for first run.")
            self.summary_label.setWordWrap(True)
            right_layout.addWidget(self.summary_label)

            self._reposition_orbs()

        @staticmethod
        def _section(text: str) -> QLabel:
            label = QLabel(text)
            label.setObjectName("SectionTitle")
            return label

        @staticmethod
        def _path_row(edit: QLineEdit, browse_cb) -> QHBoxLayout:
            row = QHBoxLayout()
            row.setSpacing(8)
            row.addWidget(edit, 1)
            btn = QPushButton("Browse")
            btn.clicked.connect(browse_cb)
            row.addWidget(btn)
            return row

        @staticmethod
        def _setup_orb(orb: QLabel, size: int, color: str) -> None:
            orb.resize(size, size)
            orb.setStyleSheet(f"background: {color}; border-radius: {size // 2}px;")
            orb.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            blur = QGraphicsBlurEffect(orb)
            blur.setBlurRadius(size * 0.25)
            orb.setGraphicsEffect(blur)
            orb.lower()

        def _reposition_orbs(self) -> None:
            w = self.centralWidget().width()
            h = self.centralWidget().height()
            self._orb_a.move(int(w * 0.68), int(h * 0.02))
            self._orb_b.move(int(w * -0.15), int(h * 0.52))
            self._orb_c.move(int(w * 0.42), int(h * 0.74))

        def _play_intro_animation(self) -> None:
            anim = QPropertyAnimation(self, b"windowOpacity")
            anim.setDuration(360)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setEasingCurve(QEasingCurve.OutCubic)
            anim.start()
            self._fade_animation = anim

        def resizeEvent(self, event) -> None:  # noqa: N802
            super().resizeEvent(event)
            self._reposition_orbs()

        def closeEvent(self, event) -> None:  # noqa: N802
            if self._thread and self._thread.isRunning():
                QMessageBox.warning(self, "Task In Progress", "Wait for the current task to finish.")
                event.ignore()
                return
            super().closeEvent(event)

        def _set_status(self, state: str, text: str) -> None:
            colors = {
                "idle": "rgba(122, 138, 168, 0.42)",
                "running": "rgba(70, 132, 255, 0.54)",
                "success": "rgba(43, 181, 120, 0.52)",
                "failed": "rgba(223, 82, 82, 0.56)",
                "warning": "rgba(204, 140, 58, 0.55)",
            }
            self.status_label.setText(text)
            self.status_label.setStyleSheet(
                "QLabel#StatusBadge {"
                f"background: {colors.get(state, colors['idle'])};"
                "border: 1px solid rgba(255, 255, 255, 0.32);"
                "}"
            )

        def _set_controls_enabled(self, enabled: bool) -> None:
            self.run_button.setEnabled(enabled)
            self.integrity_button.setEnabled(enabled)
            self.update_button.setEnabled(enabled)
            self.install_feature_button.setEnabled(enabled)
            self.sample_button.setEnabled(enabled)

        def _append_log(self, line: str) -> None:
            self.log_view.appendPlainText(line)
            bar = self.log_view.verticalScrollBar()
            bar.setValue(bar.maximum())

        def _browse_work_dir(self) -> None:
            selected = QFileDialog.getExistingDirectory(self, "Select working directory", self.work_dir_edit.text())
            if selected:
                self.work_dir_edit.setText(selected)

        def _browse_config_file(self) -> None:
            selected, _ = QFileDialog.getOpenFileName(
                self,
                "Select config file",
                self.work_dir_edit.text(),
                "JSON Files (*.json);;All Files (*)",
            )
            if selected:
                self.config_edit.setText(selected)

        def _browse_users_file(self) -> None:
            selected, _ = QFileDialog.getOpenFileName(
                self,
                "Select users file",
                self.work_dir_edit.text(),
                "JSON Files (*.json);;All Files (*)",
            )
            if selected:
                self.users_edit.setText(selected)

        def _open_relative_folder(self, relative: str) -> None:
            work_dir = Path(self.work_dir_edit.text().strip() or ".")
            folder = (work_dir / relative).resolve()
            folder.mkdir(parents=True, exist_ok=True)
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

        def _get_work_dir(self) -> Path:
            return Path(self.work_dir_edit.text().strip() or ".").expanduser().resolve()

        def _default_config_path(self, work_dir: Path) -> Path:
            local_cfg = work_dir / "config.json"
            if local_cfg.exists():
                return local_cfg
            return Path(__file__).resolve().with_name("config.json")

        def _default_users_path(self, work_dir: Path) -> Path:
            local_users = work_dir / "users.json"
            if local_users.exists():
                return local_users
            return Path(__file__).resolve().with_name("users.json")

        def _integrity_report(self, create_missing_dirs: bool = True) -> Tuple[List[str], List[str], List[str]]:
            errors: List[str] = []
            warnings: List[str] = []
            info: List[str] = []

            work_dir = self._get_work_dir()
            if not work_dir.exists():
                errors.append(f"Working directory does not exist: {work_dir}")
                return errors, warnings, info

            input_dir = work_dir / "input_data"
            output_dir = work_dir / "output_data"
            logs_dir = work_dir / "logs"

            for folder in [input_dir, output_dir, logs_dir]:
                if not folder.exists():
                    if create_missing_dirs:
                        folder.mkdir(parents=True, exist_ok=True)
                        info.append(f"Created missing directory: {folder}")
                    else:
                        errors.append(f"Missing directory: {folder}")

            xlsx_files = sorted(
                [
                    p.name for p in input_dir.glob("*.xlsx")
                    if p.is_file() and not p.name.startswith("~$")
                ]
            )
            if not xlsx_files:
                warnings.append("No input Excel files found in input_data/.")
            else:
                info.append(f"Input files detected: {len(xlsx_files)}")

            config_candidate = self.config_edit.text().strip()
            config_path = Path(config_candidate).expanduser() if config_candidate else self._default_config_path(work_dir)
            if not config_path.exists():
                errors.append(f"Config file missing: {config_path}")
            else:
                try:
                    load_config(str(config_path))
                    info.append(f"Config validated: {config_path.name}")
                except Exception as exc:  # noqa: BLE001
                    errors.append(f"Config invalid: {exc}")

            users_candidate = self.users_edit.text().strip()
            users_path = Path(users_candidate).expanduser() if users_candidate else self._default_users_path(work_dir)
            if not users_path.exists():
                errors.append(f"Users file missing: {users_path}")
            else:
                try:
                    users = load_users(str(users_path))
                    if not users:
                        warnings.append("Users list is empty.")
                    info.append(f"Users loaded: {len(users)}")
                except Exception as exc:  # noqa: BLE001
                    errors.append(f"Users file invalid: {exc}")

            for module_name in ["pandas", "openpyxl"]:
                if importlib.util.find_spec(module_name) is None:
                    errors.append(f"Dependency missing: {module_name}")

            # Validate writable output/log paths.
            for folder in [output_dir, logs_dir]:
                probe = folder / ".excel_bot_write_probe.tmp"
                try:
                    probe.write_text("ok", encoding="utf-8")
                    probe.unlink(missing_ok=True)
                except Exception as exc:  # noqa: BLE001
                    errors.append(f"Cannot write to {folder}: {exc}")

            return errors, warnings, info

        def _verify_integrity(self) -> None:
            errors, warnings, info = self._integrity_report(create_missing_dirs=True)
            self.log_view.clear()
            self._append_log("=== Integrity Check ===")
            for line in info:
                self._append_log(f"[INFO] {line}")
            for line in warnings:
                self._append_log(f"[WARN] {line}")
            for line in errors:
                self._append_log(f"[ERROR] {line}")

            if errors:
                self._set_status("failed", "Integrity Failed")
                self.summary_label.setText(
                    "Summary: integrity check failed. Resolve errors shown in console before running."
                )
                QMessageBox.critical(self, "Integrity Failed", "\n".join(errors))
                return

            if warnings:
                self._set_status("warning", "Ready With Warnings")
                self.summary_label.setText(
                    "Summary: integrity passed with warnings. You can run, but review warnings."
                )
                QMessageBox.information(self, "Integrity Check", "\n".join(warnings))
                return

            self._set_status("success", "Integrity OK")
            self.summary_label.setText("Summary: integrity check passed. System is ready.")
            QMessageBox.information(self, "Integrity Check", "All checks passed.")

        def _generate_sample_data(self) -> None:
            try:
                import pandas as pd
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Missing Dependency", f"pandas is required:\n{exc}")
                return

            work_dir = self._get_work_dir()
            input_dir = work_dir / "input_data"
            input_dir.mkdir(parents=True, exist_ok=True)
            out_path = input_dir / f"sales_data_{date.today().isoformat()}.xlsx"

            rows = [
                [1001, "Acme Corp", "Widget A", "Widgets", 10, 25.0, "Completed", "North"],
                [1002, "Beta LLC", "Widget B", "Widgets", 7, 25.0, "Completed", "North"],
                [1003, "Gamma SA", "Gadget X", "Gadgets", 4, 45.0, "Completed", "South"],
                [1004, "Helix PLC", "Service Plan", "Services", 1, 150.0, "Completed", "West"],
            ]
            columns = ["OrderID", "ClientName", "Product", "Category", "Quantity", "UnitPrice", "Status", "Region"]
            df = pd.DataFrame(rows, columns=columns)
            df.to_excel(out_path, index=False)
            self._append_log(f"[INFO] Sample file created: {out_path}")
            self.summary_label.setText(f"Summary: sample data generated at {out_path.name}.")
            QMessageBox.information(self, "Sample Data", f"Sample input created:\n{out_path}")

        def _start_task(
            self,
            task_name: str,
            command: List[str],
            work_dir: Path,
            env: Optional[Dict[str, str]] = None,
            clear_log: bool = False,
        ) -> None:
            if self._thread and self._thread.isRunning():
                return

            if clear_log:
                self.log_view.clear()

            self._set_status("running", "Running")
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)
            self._set_controls_enabled(False)

            self._thread = QThread(self)
            self._worker = CommandWorker(command, work_dir, task_name, env)
            self._worker.moveToThread(self._thread)

            self._thread.started.connect(self._worker.run)
            self._worker.line_emitted.connect(self._append_log)
            self._worker.finished.connect(self._on_task_finished)
            self._worker.finished.connect(self._thread.quit)
            self._worker.finished.connect(self._worker.deleteLater)
            self._thread.finished.connect(self._thread.deleteLater)

            self._thread.start()

        def _start_run(self) -> None:
            errors, warnings, _ = self._integrity_report(create_missing_dirs=True)
            if errors:
                QMessageBox.critical(
                    self,
                    "Cannot Run",
                    "Integrity checks failed. Fix the issues first.\n\n" + "\n".join(errors),
                )
                self._set_status("failed", "Integrity Failed")
                self.summary_label.setText("Summary: run blocked by integrity errors.")
                return

            if warnings:
                proceed = QMessageBox.question(
                    self,
                    "Warnings Detected",
                    "Warnings were detected:\n\n"
                    + "\n".join(warnings)
                    + "\n\nContinue anyway?",
                )
                if proceed != QMessageBox.Yes:
                    return

            work_dir = self._get_work_dir()
            config_path = self.config_edit.text().strip()
            users_path = self.users_edit.text().strip()

            env: Dict[str, str] = {}
            if config_path:
                env["EXCEL_BOT_CONFIG"] = str(Path(config_path).expanduser())
            if users_path:
                env["EXCEL_BOT_USERS"] = str(Path(users_path).expanduser())

            dry_run = "true" if self.mode_combo.currentIndex() == 0 else "false"
            open_files = self.open_files_check.isChecked()
            cmd = [sys.executable, "-m", "excel_bot.run_bot", "--dry-run", dry_run]
            if not open_files:
                cmd.append("--headless")

            self.summary_label.setText("Summary: running pipeline...")
            self._start_task("run", cmd, work_dir, env=env, clear_log=True)

        def _update_app(self) -> None:
            source = self.update_source_edit.text().strip() or DEFAULT_UPDATE_SOURCE
            cmd = [sys.executable, "-m", "pip", "install", "--upgrade", source]
            self.summary_label.setText("Summary: updating app package...")
            self._start_task("update", cmd, self._get_work_dir(), clear_log=False)

        def _install_feature(self) -> None:
            feature_name = self.feature_combo.currentText()
            package = FEATURE_PACKAGES.get(feature_name)
            if not package:
                QMessageBox.warning(self, "Invalid Feature", "No package mapping found for selected feature.")
                return
            cmd = [sys.executable, "-m", "pip", "install", "--upgrade", package]
            self.summary_label.setText(f"Summary: installing feature '{feature_name}'...")
            self._start_task(f"feature:{feature_name}", cmd, self._get_work_dir(), clear_log=False)

        def _on_task_finished(self, return_code: int, task_name: str) -> None:
            self.progress.setVisible(False)
            self._set_controls_enabled(True)
            self._thread = None
            self._worker = None

            if return_code == 0:
                self._set_status("success", "Completed")
                if task_name == "run":
                    self.summary_label.setText(
                        "Summary: run completed successfully.\n"
                        "Check output_data/summary_report.xlsx and logs/events.jsonl."
                    )
                elif task_name == "update":
                    self.summary_label.setText(
                        "Summary: app update completed. Restart the app to use the latest version."
                    )
                elif task_name.startswith("feature:"):
                    self.summary_label.setText(
                        f"Summary: {task_name.split(':', 1)[1]} installed successfully."
                    )
                else:
                    self.summary_label.setText("Summary: task completed successfully.")
            else:
                self._set_status("failed", f"Failed ({return_code})")
                self.summary_label.setText(
                    f"Summary: task '{task_name}' failed with exit code {return_code}. "
                    "Review console output above."
                )


def _run_gui() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = BotWindow()
    win.show()
    return app.exec()


def main() -> int:
    if _GUI_IMPORT_ERROR is not None:
        print("PySide6 is required for the desktop UI.")
        print("Install it with: pip install \"excel_bot[gui]\"")
        return 1
    return _run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
