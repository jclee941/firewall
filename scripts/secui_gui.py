from __future__ import annotations

import argparse
import sys
import traceback
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from openpyxl.utils.exceptions import InvalidFileException

from firewall_policy.gui_export import ExportFormat, ExportRequest, SourceMode, run_export_request
from firewall_policy.request_parser import RequestParseError

DEFAULT_WORKBOOK = ROOT / "dist" / "firewall-policy-automation.xlsm"
DEFAULT_FOLDER = ROOT / "request-folder"
DEFAULT_TEXT = ROOT / "dist" / "secui-cli.txt"
DEFAULT_XLSX = ROOT / "dist" / "secui-cli.xlsx"


def main() -> int:
    args = _parse_args()
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv[:1])
    window = SecuiCliWindow()
    if args.smoke:
        print(f"SECUI_CLI_GUI_SMOKE_OK {len(window.windowTitle())}")
        return 0
    window.show()
    return app.exec()


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run native SECUI CLI GUI.")
    parser.add_argument("--smoke", action="store_true", help="Create the native GUI window and exit.")
    return parser.parse_args()


class SecuiCliWindow:
    def __init__(self) -> None:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import (
            QComboBox,
            QFileDialog,
            QGridLayout,
            QGroupBox,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QMessageBox,
            QPushButton,
            QRadioButton,
            QTextEdit,
            QVBoxLayout,
            QWidget,
        )

        self._qt = Qt
        self._file_dialog = QFileDialog
        self._message_box = QMessageBox
        self.widget = QWidget()
        self.widget.setWindowTitle("SECUI CLI 생성기")
        self.widget.setMinimumSize(760, 430)

        self.workbook = QLineEdit(str(DEFAULT_WORKBOOK))
        self.request_folder = QLineEdit(str(DEFAULT_FOLDER))
        self.config_workbook = QLineEdit()
        self.parse_sheet = QLineEdit()
        self.text_output = QLineEdit(str(DEFAULT_TEXT))
        self.xlsx_output = QLineEdit(str(DEFAULT_XLSX))
        self.source_workbook = QRadioButton("Workbook")
        self.source_folder = QRadioButton("Request Folder")
        self.source_workbook.setChecked(True)
        self.output_format = QComboBox()
        self.output_format.addItems(["text", "xlsx", "both"])
        self.output_format.setCurrentText("both")
        self.status = QTextEdit("대기")
        self.status.setReadOnly(True)
        self.status.setFixedHeight(58)

        layout = QVBoxLayout(self.widget)
        title = QLabel("SECUI CLI 생성기")
        title.setStyleSheet("font-size:22px;font-weight:700;color:#1b2430")
        layout.addWidget(title)
        layout.addWidget(self._input_group())
        layout.addWidget(self._output_group())
        layout.addWidget(self.status)
        export = QPushButton("Export")
        export.clicked.connect(self._export)
        layout.addWidget(export, alignment=Qt.AlignmentFlag.AlignRight)

    def _input_group(self):
        from PySide6.QtWidgets import QGridLayout, QGroupBox, QHBoxLayout, QPushButton, QWidget

        group = QGroupBox("입력")
        grid = QGridLayout(group)
        source_row = QWidget()
        source_layout = QHBoxLayout(source_row)
        source_layout.setContentsMargins(0, 0, 0, 0)
        source_layout.addWidget(self.source_workbook)
        source_layout.addWidget(self.source_folder)
        source_layout.addStretch(1)
        grid.addWidget(source_row, 0, 0, 1, 3)
        self._path_row(grid, 1, "Workbook", self.workbook, self._pick_workbook)
        self._path_row(grid, 2, "요청 폴더", self.request_folder, self._pick_request_folder)
        self._path_row(grid, 3, "설정 Workbook", self.config_workbook, self._pick_config_workbook)
        grid.addWidget(self._label("파싱 시트"), 4, 0)
        grid.addWidget(self.parse_sheet, 4, 1, 1, 2)
        return group

    def _output_group(self):
        from PySide6.QtWidgets import QGridLayout, QGroupBox

        group = QGroupBox("Export")
        grid = QGridLayout(group)
        grid.addWidget(self._label("형식"), 0, 0)
        grid.addWidget(self.output_format, 0, 1, 1, 2)
        self._path_row(grid, 1, "Text 파일", self.text_output, self._pick_text_output)
        self._path_row(grid, 2, "XLSX 파일", self.xlsx_output, self._pick_xlsx_output)
        return group

    def _path_row(self, grid, row: int, text: str, field, picker) -> None:
        from PySide6.QtWidgets import QPushButton

        grid.addWidget(self._label(text), row, 0)
        grid.addWidget(field, row, 1)
        button = QPushButton("선택")
        button.clicked.connect(picker)
        grid.addWidget(button, row, 2)

    def _label(self, text: str):
        from PySide6.QtWidgets import QLabel

        label = QLabel(text)
        label.setMinimumWidth(120)
        return label

    def windowTitle(self) -> str:
        return self.widget.windowTitle()

    def show(self) -> None:
        self.widget.show()

    def _pick_workbook(self) -> None:
        self._set_path(self.workbook, self._file_dialog.getOpenFileName(self.widget, "Workbook", str(ROOT), "Excel (*.xlsx *.xlsm);;All (*)")[0])

    def _pick_config_workbook(self) -> None:
        self._set_path(self.config_workbook, self._file_dialog.getOpenFileName(self.widget, "설정 Workbook", str(ROOT), "Excel (*.xlsx *.xlsm);;All (*)")[0])

    def _pick_request_folder(self) -> None:
        self._set_path(self.request_folder, self._file_dialog.getExistingDirectory(self.widget, "요청 폴더", str(ROOT)))

    def _pick_text_output(self) -> None:
        self._set_path(self.text_output, self._file_dialog.getSaveFileName(self.widget, "Text 파일", str(DEFAULT_TEXT), "Text (*.txt);;All (*)")[0])

    def _pick_xlsx_output(self) -> None:
        self._set_path(self.xlsx_output, self._file_dialog.getSaveFileName(self.widget, "XLSX 파일", str(DEFAULT_XLSX), "Excel (*.xlsx);;All (*)")[0])

    def _set_path(self, field, value: str) -> None:
        if value:
            field.setText(value)

    def _export(self) -> None:
        try:
            run_export_request(self._request())
        except ValueError as exc:
            self._fail(str(exc))
            return
        except FileNotFoundError as exc:
            self._fail(f"file not found: {exc.filename or exc}")
            return
        except (RequestParseError, zipfile.BadZipFile, InvalidFileException, KeyError, ET.ParseError) as exc:
            self._fail(f"workbook/folder is unreadable: {exc}")
            return
        except OSError as exc:
            self._fail(f"I/O error: {exc}")
            return
        except Exception as exc:  # last-resort safety net so the window never crashes
            traceback.print_exc()
            self._fail(f"unexpected error: {type(exc).__name__}: {exc}")
            return
        self.status.setPlainText("완료")
        self._message_box.information(self.widget, "Export 완료", "export 완료")

    def _request(self) -> ExportRequest:
        return ExportRequest(
            source_mode=_source_mode(self.source_workbook.isChecked()),
            output_format=_export_format(self.output_format.currentText()),
            workbook=_path_or_none(self.workbook.text()),
            request_folder=_path_or_none(self.request_folder.text()),
            config_workbook=_path_or_none(self.config_workbook.text()),
            parse_sheet=self.parse_sheet.text().strip(),
            text_output=_path_or_none(self.text_output.text()),
            xlsx_output=_path_or_none(self.xlsx_output.text()),
        )

    def _fail(self, message: str) -> None:
        self.status.setPlainText(message)
        self._message_box.critical(self.widget, "Export 실패", message)


def _source_mode(workbook_checked: bool) -> SourceMode:
    return "workbook" if workbook_checked else "request_folder"


def _export_format(value: str) -> ExportFormat:
    match value:
        case "text":
            return "text"
        case "xlsx":
            return "xlsx"
        case "both":
            return "both"
        case _:
            raise ValueError(f"unknown export format: {value}")


def _path_or_none(value: str) -> Path | None:
    text = value.strip()
    return Path(text) if text else None


if __name__ == "__main__":
    raise SystemExit(main())
