from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CI_WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
RELEASE_WORKFLOW = ROOT / ".github" / "workflows" / "release.yml"


def _literal_run_blocks(src: str) -> list[str]:
    lines = src.splitlines()
    blocks: list[str] = []
    for index, line in enumerate(lines):
        if line.strip() != "run: |":
            continue
        indent = len(line) - len(line.lstrip())
        block: list[str] = []
        for body_line in lines[index + 1 :]:
            if body_line.strip() and len(body_line) - len(body_line.lstrip()) <= indent:
                break
            block.append(body_line)
        blocks.append("\n".join(block))
    return blocks


def test_ci_build_job_is_read_only_and_auto_tag_job_dispatches_release() -> None:
    src = CI_WORKFLOW.read_text(encoding="utf-8")

    assert "permissions:\n  contents: read" in src
    assert "persist-credentials: false" in src
    assert "create-release-tag:" in src
    assert "needs: build-test" in src
    assert "contents: write" in src
    assert "actions: write" in src
    assert "git push origin \"$next\"" in src
    assert "gh workflow run release.yml --ref \"$next\" -f tag=\"$next\"" in src
    assert "git tag --points-at HEAD --list" in src
    assert "grep -E '^v[0-9]+\\.[0-9]+\\.[0-9]+$'" in src
    assert "--list 'v[0-9]*.[0-9]*.[0-9]*'" not in src


def test_release_workflow_dispatch_checks_out_and_validates_exact_tag() -> None:
    src = RELEASE_WORKFLOW.read_text(encoding="utf-8")

    assert "ref: ${{ github.event_name == 'workflow_dispatch' && github.event.inputs.tag || github.ref }}" in src
    assert src.count("name: Validate release tag") == 2
    assert "^[v][0-9]+\\.[0-9]+\\.[0-9]+$" in src
    assert "git rev-list -n 1 \"$TAG\"" in src
    assert "git rev-parse HEAD" in src
    assert "tag $TAG does not match checked-out commit" in src


def test_release_workflow_does_not_interpolate_manual_tag_inside_shell() -> None:
    src = RELEASE_WORKFLOW.read_text(encoding="utf-8")
    run_blocks = "\n".join(_literal_run_blocks(src))

    assert "${{ github.event.inputs.tag }}" not in run_blocks
    assert src.count("RELEASE_TAG_INPUT: ${{ github.event.inputs.tag }}") == 2
    assert src.count('TAG="$RELEASE_TAG_INPUT"') == 2


def test_release_workflow_limits_write_scope_and_does_not_clobber_assets() -> None:
    src = RELEASE_WORKFLOW.read_text(encoding="utf-8")

    assert "permissions:\n  contents: read" in src
    assert src.count("permissions:\n      contents: write") == 2
    assert "--clobber" not in src
