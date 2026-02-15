"""Tests for the CLI module."""

import tempfile
from pathlib import Path

from typer.testing import CliRunner

from docxfix.cli import app

runner = CliRunner()


def test_info_command():
    """Test the info command."""
    result = runner.invoke(app, ["info"])
    assert result.exit_code == 0
    assert "docxfix version" in result.stdout
    assert "0.1.0" in result.stdout


def test_create_command_basic():
    """Test the create command with basic arguments."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.docx"
        result = runner.invoke(app, ["create", str(output_path)])
        assert result.exit_code == 0
        assert "created successfully" in result.stdout
        assert output_path.exists()


def test_create_command_verbose():
    """Test the create command with verbose flag."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.docx"
        result = runner.invoke(app, ["create", str(output_path), "--verbose"])
        assert result.exit_code == 0
        assert "Creating docx fixture" in result.stdout
        assert "Validating document" in result.stdout
        assert "Validation passed" in result.stdout


def test_create_command_with_template():
    """Test the create command with template option."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.docx"
        result = runner.invoke(
            app,
            ["create", str(output_path), "--template", "template.docx", "-v"],
        )
        assert result.exit_code == 0
        assert "Using template: template.docx" in result.stdout


def test_create_command_no_validate():
    """Test the create command without validation."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.docx"
        result = runner.invoke(
            app, ["create", str(output_path), "--no-validate"]
        )
        assert result.exit_code == 0
        assert "created successfully" in result.stdout
        # Should still create the file
        assert output_path.exists()
