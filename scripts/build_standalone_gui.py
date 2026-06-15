from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "dist" / "secui-gui-standalone.zip"


def main() -> int:
    args = _parse_args()
    output = Path(args.output)
    if not args.no_rebuild:
        _run([sys.executable, str(ROOT / "scripts" / "build_xlsm.py")])
        _run([sys.executable, str(ROOT / "scripts" / "make_request_folder.py")])
    stage = output.with_suffix("")
    if stage.exists():
        shutil.rmtree(stage)
    exe = _build_executable(stage)
    _stage_runtime(stage, exe)
    _zip_dir(stage, output)
    print(f"Built {output}")
    return 0


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build no-Python SECUI GUI standalone bundle.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--no-rebuild", action="store_true")
    return parser.parse_args()


def _run(command: list[str]) -> None:
    subprocess.run(command, cwd=ROOT, check=True)


def _build_executable(stage: Path) -> Path:
    dist_path = stage / "_pyinstaller_dist"
    work_path = stage / "_pyinstaller_build"
    spec_path = stage / "_pyinstaller_spec"
    _run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--name",
            "secui-gui",
            "--distpath",
            str(dist_path),
            "--workpath",
            str(work_path),
            "--specpath",
            str(spec_path),
            str(ROOT / "scripts" / "secui_gui.py"),
        ]
    )
    candidates = sorted(path for path in dist_path.iterdir() if path.name.startswith("secui-gui"))
    if not candidates:
        raise FileNotFoundError(str(dist_path / "secui-gui"))
    return candidates[0]


def _stage_runtime(stage: Path, exe: Path) -> None:
    runtime = stage / "secui-gui"
    (runtime / "dist").mkdir(parents=True)
    shutil.copy2(exe, runtime / exe.name)
    shutil.copy2(ROOT / "dist" / "firewall-policy-automation.xlsm", runtime / "dist" / "firewall-policy-automation.xlsm")
    shutil.copytree(ROOT / "request-folder", runtime / "request-folder")
    (runtime / "README.md").write_text("\n".join(_readme_lines(exe.name)), encoding="utf-8", newline="\n")
    shutil.rmtree(stage / "_pyinstaller_dist")
    shutil.rmtree(stage / "_pyinstaller_build")
    shutil.rmtree(stage / "_pyinstaller_spec")


def _readme_lines(exe_name: str) -> list[str]:
    return [
        "# SECUI GUI Standalone",
        "",
        "Python 설치 없이 실행하는 SECUI CLI native GUI 번들입니다.",
        "",
        "## 실행",
        f"- Windows: `{exe_name}` 더블클릭",
        f"- Linux/macOS: `./{exe_name}` 실행",
        "",
        "실행 후 GUI 창에서 workbook/request-folder와 export 형식을 선택하고 Export를 누르세요.",
        "`dist/firewall-policy-automation.xlsm`과 `request-folder/` 샘플이 함께 포함되어 있습니다.",
    ]


def _zip_dir(source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted((source / "secui-gui").rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source))


if __name__ == "__main__":
    raise SystemExit(main())
