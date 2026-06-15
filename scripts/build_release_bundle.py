from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "dist" / "secui-cli-gui-release.zip"
GUI_REQUIREMENTS = ROOT / "requirements-gui.txt"


def main() -> int:
    args = _parse_args()
    output = Path(args.output)
    if not args.no_rebuild:
        _run([sys.executable, str(ROOT / "scripts" / "build_xlsm.py")])
        _run([sys.executable, str(ROOT / "scripts" / "make_request_folder.py")])
    bundle_root = output.with_suffix("")
    if bundle_root.exists():
        shutil.rmtree(bundle_root)
    _stage_bundle(bundle_root)
    if not args.skip_wheel_download:
        _download_wheels(bundle_root / "vendor" / "wheels")
    _zip_dir(bundle_root, output)
    print(f"Built {output}")
    return 0


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build SECUI Python GUI/CLI release bundle with bundled wheels.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output .zip path.")
    parser.add_argument("--no-rebuild", action="store_true", help="Skip workbook/request-folder rebuild before bundling.")
    parser.add_argument("--skip-wheel-download", action="store_true", help="Create bundle layout without downloading wheels.")
    return parser.parse_args()


def _run(command: list[str]) -> None:
    subprocess.run(command, cwd=ROOT, check=True)


def _stage_bundle(bundle_root: Path) -> None:
    app = bundle_root / "app"
    (app / "dist").mkdir(parents=True)
    (bundle_root / "vendor" / "wheels").mkdir(parents=True)
    shutil.copytree(ROOT / "firewall_policy", app / "firewall_policy", ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copytree(ROOT / "scripts", app / "scripts", ignore=shutil.ignore_patterns("__pycache__", "AGENTS.md"))
    shutil.copytree(ROOT / "request-folder", app / "request-folder")
    shutil.copy2(ROOT / "dist" / "firewall-policy-automation.xlsm", app / "dist" / "firewall-policy-automation.xlsm")
    shutil.copy2(ROOT / "requirements.txt", bundle_root / "requirements.txt")
    shutil.copy2(GUI_REQUIREMENTS, bundle_root / "requirements-gui.txt")
    if (ROOT / "dist" / "README.txt").exists():
        shutil.copy2(ROOT / "dist" / "README.txt", bundle_root / "EXCEL-README.txt")
    _write_text(bundle_root / "README.md", _readme())
    _write_text(bundle_root / "install_offline.sh", _install_sh())
    _write_text(bundle_root / "run_gui.sh", _run_gui_sh())
    _write_text(bundle_root / "run_cli_export_example.sh", _run_cli_sh())
    _write_text(bundle_root / "install_offline.bat", _install_bat())
    _write_text(bundle_root / "run_gui.bat", _run_gui_bat())
    _write_text(bundle_root / "run_cli_export_example.bat", _run_cli_bat())
    for script in ("install_offline.sh", "run_gui.sh", "run_cli_export_example.sh"):
        (bundle_root / script).chmod(0o755)


def _download_wheels(wheel_dir: Path) -> None:
    _run([sys.executable, "-m", "pip", "download", "--dest", str(wheel_dir), "-r", str(GUI_REQUIREMENTS)])


def _zip_dir(source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(source.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source.parent))


def _write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8", newline="\n")


def _readme() -> str:
    return """# SECUI CLI GUI Release

이 zip은 VBA 실행 없이 SECUI CLI를 생성하는 Python native GUI/CLI 번들입니다.

## 포함 파일
- `app/scripts/secui_gui.py`: native GUI 실행 파일
- `app/scripts/secui_cli.py`: CLI export 실행 파일
- `app/firewall_policy/`: Excel 파싱 및 export 지원 코드
- `app/dist/firewall-policy-automation.xlsm`: 샘플/설정 workbook
- `app/request-folder/`: 샘플 요청 폴더
- `vendor/wheels/`: `requirements-gui.txt` 기준 pip 의존성 wheel

## 설치
Python 3.12가 설치된 PC에서 아래 파일을 실행하세요.

Linux/macOS:
```bash
./install_offline.sh
./run_gui.sh
```

Windows:
```bat
install_offline.bat
run_gui.bat
```
"""


def _install_sh() -> str:
    return """#!/usr/bin/env sh
set -eu
python3 -m venv .venv
.venv/bin/python -m pip install --no-index --find-links vendor/wheels -r requirements-gui.txt
"""


def _run_gui_sh() -> str:
    return """#!/usr/bin/env sh
set -eu
.venv/bin/python app/scripts/secui_gui.py
"""


def _run_cli_sh() -> str:
    return """#!/usr/bin/env sh
set -eu
.venv/bin/python app/scripts/secui_cli.py export --workbook app/dist/firewall-policy-automation.xlsm --format both --output secui-cli.txt --output-xlsx secui-cli.xlsx
"""


def _install_bat() -> str:
    return """@echo off
py -3.12 -m venv .venv
.venv\\Scripts\\python.exe -m pip install --no-index --find-links vendor\\wheels -r requirements-gui.txt
"""


def _run_gui_bat() -> str:
    return """@echo off
.venv\\Scripts\\python.exe app\\scripts\\secui_gui.py
"""


def _run_cli_bat() -> str:
    return """@echo off
.venv\\Scripts\\python.exe app\\scripts\\secui_cli.py export --workbook app\\dist\\firewall-policy-automation.xlsm --format both --output secui-cli.txt --output-xlsx secui-cli.xlsx
"""


if __name__ == "__main__":
    raise SystemExit(main())
