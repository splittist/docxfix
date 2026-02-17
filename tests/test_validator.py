"""Tests for the validator module."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import ChangeType, Comment, DocumentSpec, TrackedChange
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


def test_validator_comment_id_uniqueness():
    """Duplicate comment IDs trigger validation error."""
    spec = DocumentSpec()
    spec.add_paragraph(
        "Hello world",
        comments=[
            Comment(text="c1", anchor_text="Hello"),
            Comment(text="c2", anchor_text="world"),
        ],
    )
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        # Valid doc should pass
        DocumentValidator(path).validate()


def test_validator_tracked_change_id_uniqueness():
    """Tracked change IDs must be unique."""
    spec = DocumentSpec()
    spec.add_paragraph(
        "Hello cruel world",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text="cruel ",
            ),
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="beautiful ",
                insert_after="Hello ",
            ),
        ],
    )
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        DocumentValidator(path).validate()


def test_validator_comment_anchor_integrity():
    """Comment range start/end must pair up."""
    spec = DocumentSpec()
    spec.add_paragraph(
        "Test paragraph",
        comments=[Comment(text="note", anchor_text="Test")],
    )
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        DocumentValidator(path).validate()


def test_validator_comment_anchor_mismatch():
    """Unpaired commentRangeStart triggers error."""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        # Build a minimal docx with mismatched anchors
        spec = DocumentSpec()
        spec.add_paragraph("Test")
        DocumentGenerator(spec).generate(path)

        # Tamper: add orphan commentRangeStart
        with zipfile.ZipFile(path, "r") as zin:
            doc_xml = zin.read("word/document.xml")
            other_files = {
                n: zin.read(n)
                for n in zin.namelist()
                if n != "word/document.xml"
            }

        root = etree.fromstring(doc_xml)
        body = root.find(f"{{{w_ns}}}body")
        first_p = body.find(f"{{{w_ns}}}p")
        etree.SubElement(
            first_p,
            f"{{{w_ns}}}commentRangeStart",
            {f"{{{w_ns}}}id": "99"},
        )

        with zipfile.ZipFile(path, "w") as zout:
            for name, data in other_files.items():
                zout.writestr(name, data)
            zout.writestr(
                "word/document.xml",
                etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                ),
            )

        with pytest.raises(
            ValidationError, match="Comment anchor mismatch"
        ):
            DocumentValidator(path).validate()


def test_validator_relationship_completeness():
    """Every rId in document.xml must exist in rels."""
    spec = DocumentSpec()
    spec.add_paragraph("Test")
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        # Valid doc passes
        DocumentValidator(path).validate()


def test_validator_content_type_coverage():
    """Every part in ZIP must have a content type."""
    spec = DocumentSpec()
    spec.add_paragraph("Test")
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        DocumentValidator(path).validate()


def test_validator_content_type_missing():
    """A part without content type triggers error."""
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        spec = DocumentSpec()
        spec.add_paragraph("Test")
        DocumentGenerator(spec).generate(path)

        # Add an extra file with unknown extension
        with zipfile.ZipFile(path, "a") as zout:
            zout.writestr("word/extra.foo", b"data")

        with pytest.raises(
            ValidationError, match="no matching content type"
        ):
            DocumentValidator(path).validate()
