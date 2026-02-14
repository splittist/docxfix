"""Tests for the generator module."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import ChangeType, DocumentSpec, TrackedChange


def test_generator_creates_zip_file():
    """Test that generator creates a valid ZIP file."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        assert output_path.exists()
        assert zipfile.is_zipfile(output_path)


def test_generator_includes_required_files():
    """Test that generated docx includes required files."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            files = docx_zip.namelist()
            assert "[Content_Types].xml" in files
            assert "_rels/.rels" in files
            assert "word/document.xml" in files
            assert "word/_rels/document.xml.rels" in files


def test_generator_creates_valid_xml():
    """Test that generated XML files are well-formed."""
    spec = DocumentSpec()
    spec.add_paragraph("Test paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            # Parse each XML file to verify well-formedness
            for filename in docx_zip.namelist():
                if filename.endswith(".xml") or filename.endswith(".rels"):
                    content = docx_zip.read(filename)
                    # Should not raise exception
                    etree.fromstring(content)


def test_generator_simple_paragraph():
    """Test generating a document with a simple paragraph."""
    spec = DocumentSpec()
    spec.add_paragraph("Hello, World!")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            doc_xml = docx_zip.read("word/document.xml")
            root = etree.fromstring(doc_xml)

            # Find namespace
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            # Check paragraph exists
            paragraphs = root.findall(".//w:p", namespaces=ns)
            assert len(paragraphs) >= 1

            # Check text exists
            text_elements = root.findall(".//w:t", namespaces=ns)
            assert any(
                elem.text == "Hello, World!" for elem in text_elements
            )


def test_generator_multiple_paragraphs():
    """Test generating a document with multiple paragraphs."""
    spec = DocumentSpec()
    spec.add_paragraph("First paragraph")
    spec.add_paragraph("Second paragraph")
    spec.add_paragraph("Third paragraph")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            doc_xml = docx_zip.read("word/document.xml")
            root = etree.fromstring(doc_xml)

            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            paragraphs = root.findall(".//w:p", namespaces=ns)
            assert len(paragraphs) == 3


def test_generator_tracked_insertion():
    """Test generating a document with tracked insertion."""
    spec = DocumentSpec()
    change = TrackedChange(
        change_type=ChangeType.INSERTION,
        text="inserted text",
        author="Test Author",
    )
    spec.add_paragraph("Main text", tracked_changes=[change])

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            doc_xml = docx_zip.read("word/document.xml")
            root = etree.fromstring(doc_xml)

            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            # Check insertion element exists
            insertions = root.findall(".//w:ins", namespaces=ns)
            assert len(insertions) == 1

            # Check insertion attributes
            ins_elem = insertions[0]
            assert ins_elem.get(f"{{{ns['w']}}}author") == "Test Author"
            assert f"{{{ns['w']}}}id" in ins_elem.attrib
            assert f"{{{ns['w']}}}date" in ins_elem.attrib

            # Check insertion contains text
            text_in_ins = ins_elem.findall(".//w:t", namespaces=ns)
            assert len(text_in_ins) == 1
            assert text_in_ins[0].text == "inserted text"


def test_generator_tracked_deletion():
    """Test generating a document with tracked deletion."""
    spec = DocumentSpec()
    change = TrackedChange(
        change_type=ChangeType.DELETION,
        text="deleted text",
        author="Test Author",
    )
    spec.add_paragraph("Main text", tracked_changes=[change])

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            doc_xml = docx_zip.read("word/document.xml")
            root = etree.fromstring(doc_xml)

            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            # Check deletion element exists
            deletions = root.findall(".//w:del", namespaces=ns)
            assert len(deletions) == 1

            # Check deletion attributes
            del_elem = deletions[0]
            assert del_elem.get(f"{{{ns['w']}}}author") == "Test Author"

            # Check deletion contains delText
            del_text = del_elem.findall(".//w:delText", namespaces=ns)
            assert len(del_text) == 1
            assert del_text[0].text == "deleted text"


def test_generator_multiple_tracked_changes():
    """Test generating with multiple tracked changes."""
    spec = DocumentSpec()
    changes = [
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="first insert",
        ),
        TrackedChange(
            change_type=ChangeType.DELETION,
            text="deleted",
        ),
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="second insert",
        ),
    ]
    spec.add_paragraph("Main text", tracked_changes=changes)

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            doc_xml = docx_zip.read("word/document.xml")
            root = etree.fromstring(doc_xml)

            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            # Check we have 2 insertions and 1 deletion
            insertions = root.findall(".//w:ins", namespaces=ns)
            deletions = root.findall(".//w:del", namespaces=ns)

            assert len(insertions) == 2
            assert len(deletions) == 1
