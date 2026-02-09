"""Tests for the CLI module."""

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
    result = runner.invoke(app, ["create", "output.docx"])
    assert result.exit_code == 0
    assert "created successfully" in result.stdout


def test_create_command_verbose():
    """Test the create command with verbose flag."""
    result = runner.invoke(app, ["create", "output.docx", "--verbose"])
    assert result.exit_code == 0
    assert "Creating docx fixture" in result.stdout


def test_create_command_with_template():
    """Test the create command with template option."""
    result = runner.invoke(
        app, ["create", "output.docx", "--template", "template.docx", "-v"]
    )
    assert result.exit_code == 0
    assert "Using template: template.docx" in result.stdout
