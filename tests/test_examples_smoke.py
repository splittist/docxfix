"""Smoke tests for the bundled example specs and BDD example script.

These tests verify that every example in ``examples/`` can be processed
end-to-end by the CLI and that the standalone ``bdd_row_example.py`` runs
without error.  They serve as CI smoke checks for M3.4.
"""

import subprocess
import sys
from pathlib import Path

import pytest
from typer.testing import CliRunner

from docxfix.cli import app

EXAMPLES_DIR = Path(__file__).parent.parent / "examples"
runner = CliRunner()

# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

EXAMPLE_SPECS = [
    "01-simple.yaml",
    "02-tracked-changes.yaml",
    "03-comments.yaml",
    "04-legal-list-numbering.yaml",
    "05-heading-numbering.yaml",
    "06-sections.yaml",
    "07-combined.yaml",
    "08-deterministic.json",
]


# ---------------------------------------------------------------------------
# Smoke: each example spec generates a valid .docx via CLI
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("spec_filename", EXAMPLE_SPECS)
def test_example_spec_generates_valid_docx(spec_filename, tmp_path):
    """Each bundled example spec must produce a valid .docx via `create --spec`."""
    spec_file = EXAMPLES_DIR / spec_filename
    assert spec_file.exists(), f"Example spec not found: {spec_file}"

    output_file = tmp_path / "output.docx"
    result = runner.invoke(
        app,
        ["create", str(output_file), "--spec", str(spec_file)],
    )

    assert result.exit_code == 0, (
        f"CLI failed for {spec_filename}:\n{result.output}"
    )
    assert output_file.exists(), f"Output file not created for {spec_filename}"
    assert output_file.stat().st_size > 0, f"Output file is empty for {spec_filename}"


# ---------------------------------------------------------------------------
# Smoke: batch manifest generates all 8 fixtures
# ---------------------------------------------------------------------------


def test_batch_manifest_generates_all_fixtures(tmp_path):
    """The bundled batch-manifest.yaml must generate all 8 fixtures successfully."""
    manifest_file = EXAMPLES_DIR / "batch-manifest.yaml"
    assert manifest_file.exists(), "batch-manifest.yaml not found in examples/"

    out_dir = tmp_path / "fixtures"
    result = runner.invoke(
        app,
        ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)],
    )

    assert result.exit_code == 0, (
        f"Batch command failed:\n{result.output}"
    )
    assert "Success: 8" in result.output, (
        f"Expected 8 successes:\n{result.output}"
    )
    assert "Failed: 0" in result.output, (
        f"Expected 0 failures:\n{result.output}"
    )

    expected_outputs = [
        "simple.docx",
        "tracked-changes.docx",
        "comments.docx",
        "legal-list.docx",
        "heading-numbering.docx",
        "sections.docx",
        "combined.docx",
        "deterministic.docx",
    ]
    for filename in expected_outputs:
        assert (out_dir / filename).exists(), f"Missing output: {filename}"


# ---------------------------------------------------------------------------
# Smoke: bdd_row_example.py runs without error
# ---------------------------------------------------------------------------


def test_bdd_row_example_script_runs(tmp_path):
    """examples/bdd_row_example.py must run to completion with exit code 0."""
    script = EXAMPLES_DIR / "bdd_row_example.py"
    assert script.exists(), "bdd_row_example.py not found in examples/"

    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, (
        f"bdd_row_example.py failed:\nstdout: {result.stdout}\nstderr: {result.stderr}"
    )
    assert "succeeded" in result.stdout, (
        f"Expected success summary in output:\n{result.stdout}"
    )
    assert "failed" in result.stdout
    # All 5 scenario rows should succeed
    assert "5 succeeded, 0 failed" in result.stdout, (
        f"Expected all 5 to succeed:\n{result.stdout}"
    )
