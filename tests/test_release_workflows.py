from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CI_WORKFLOW = ROOT / ".github" / "workflows" / "ci.yml"
RELEASE_WORKFLOW = ROOT / ".github" / "workflows" / "release.yml"


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


def test_release_workflow_dispatch_checks_out_and_validates_exact_tag() -> None:
    src = RELEASE_WORKFLOW.read_text(encoding="utf-8")

    assert "ref: ${{ github.event_name == 'workflow_dispatch' && github.event.inputs.tag || github.ref }}" in src
    assert src.count("name: Validate release tag") == 2
    assert "^[v][0-9]+\\.[0-9]+\\.[0-9]+$" in src
    assert "git rev-list -n 1 \"$TAG\"" in src
    assert "git rev-parse HEAD" in src
    assert "tag $TAG does not match checked-out commit" in src
