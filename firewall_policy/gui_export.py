from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, assert_never
from argparse import Namespace

SourceMode = Literal["workbook", "request_folder"]
ExportFormat = Literal["text", "xlsx", "both"]


@dataclass(frozen=True, slots=True)
class ExportRequest:
    source_mode: SourceMode
    output_format: ExportFormat
    workbook: Path | None = None
    request_folder: Path | None = None
    config_workbook: Path | None = None
    parse_sheet: str = ""
    text_output: Path | None = None
    xlsx_output: Path | None = None


def build_export_command(python_executable: str, script_path: Path, request: ExportRequest) -> list[str]:
    _validate_request(request)
    command = [python_executable, str(script_path), "export"]

    match request.source_mode:
        case "workbook":
            if request.workbook is None:
                raise ValueError("workbook source requires a workbook path")
            command.extend(["--workbook", str(request.workbook)])
        case "request_folder":
            if request.request_folder is None:
                raise ValueError("request-folder source requires a folder path")
            command.extend(["--request-folder", str(request.request_folder)])
            if request.config_workbook is not None:
                command.extend(["--config-workbook", str(request.config_workbook)])
            if request.parse_sheet:
                command.extend(["--parse-sheet", request.parse_sheet])
        case unreachable:
            assert_never(unreachable)

    command.extend(["--format", request.output_format])
    match request.output_format:
        case "text":
            if request.text_output is not None:
                command.extend(["--output", str(request.text_output)])
        case "xlsx":
            if request.xlsx_output is not None:
                command.extend(["--output-xlsx", str(request.xlsx_output)])
        case "both":
            if request.text_output is not None:
                command.extend(["--output", str(request.text_output)])
            if request.xlsx_output is not None:
                command.extend(["--output-xlsx", str(request.xlsx_output)])
        case unreachable:
            assert_never(unreachable)
    return command


def run_export_request(request: ExportRequest) -> None:
    _validate_request(request)
    from scripts import secui_cli

    args = Namespace(
        workbook=str(request.workbook) if request.workbook is not None else None,
        request_folder=str(request.request_folder) if request.request_folder is not None else None,
        config_workbook=str(request.config_workbook) if request.config_workbook is not None else None,
        parse_sheet=request.parse_sheet,
        format=request.output_format,
        output=str(request.text_output) if request.text_output is not None else None,
        output_xlsx=str(request.xlsx_output) if request.xlsx_output is not None else None,
    )
    rows = secui_cli._run(args)
    secui_cli._export_rows(args, rows)


def _validate_request(request: ExportRequest) -> None:
    match request.output_format:
        case "text":
            if request.text_output is None:
                raise ValueError("text export requires a text output path")
        case "xlsx":
            if request.xlsx_output is None:
                raise ValueError("xlsx export requires an xlsx output path")
        case "both":
            if request.text_output is None or request.xlsx_output is None:
                raise ValueError("both export requires text and xlsx output paths")
        case unreachable:
            assert_never(unreachable)
