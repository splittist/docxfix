"""Tests for the validator module."""

import tempfile
import zipfile
from pathlib import Path

import pytest

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec
from docxfix.validator import DocumentValidator, ValidationError, validate_docx


def test_validator_valid_document():
    """Test validating a valid generated document."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        # Should not raise exception
        validator = DocumentValidator(output_path)
        validator.validate()


def test_validate_docx_function():
    """Test the validate_docx convenience function."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        # Should not raise exception
        validate_docx(output_path)


def test_validator_missing_file():
    """Test validator with non-existent file."""
    validator = DocumentValidator("nonexistent.docx")

    with pytest.raises(ValidationError, match="File not found"):
        validator.validate()


def test_validator_not_zip_file():
    """Test validator with non-ZIP file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        invalid_path = Path(tmpdir) / "not_a_zip.docx"
        invalid_path.write_text("This is not a ZIP file")

        validator = DocumentValidator(invalid_path)

        with pytest.raises(ValidationError, match="Not a valid ZIP file"):
            validator.validate()


def test_validator_missing_required_file():
    """Test validator with missing required file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "incomplete.docx"

        # Create ZIP with only some files
        with zipfile.ZipFile(output_path, "w") as docx_zip:
            docx_zip.writestr("[Content_Types].xml", "<Types/>")
            # Missing _rels/.rels and word/document.xml

        validator = DocumentValidator(output_path)

        with pytest.raises(ValidationError, match="Missing required file"):
            validator.validate()


def test_validator_malformed_xml():
    """Test validator with malformed XML."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "malformed.docx"

        # Create ZIP with malformed XML
        with zipfile.ZipFile(output_path, "w") as docx_zip:
            docx_zip.writestr("[Content_Types].xml", "<Types>")
            docx_zip.writestr(
                "_rels/.rels",
                '<?xml version="1.0"?><Relationships></Relationships>',
            )
            docx_zip.writestr(
                "word/document.xml",
                '<?xml version="1.0"?><Invalid><not>closed',
            )

        validator = DocumentValidator(output_path)

        with pytest.raises(ValidationError, match="XML syntax error"):
            validator.validate()


def test_validator_all_files_wellformed():
    """Test that validator checks all XML files."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        # Verify all XML files are checked
        validator = DocumentValidator(output_path)

        # This should pass without issues
        validator._validate_xml_wellformedness()
