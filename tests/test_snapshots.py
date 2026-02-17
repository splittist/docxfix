"""Snapshot tests for generated XML parts using syrupy."""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    Comment,
    CommentReply,
    DocumentSpec,
    NumberedParagraph,
)


def _extract_xml(docx_path: Path, part_name: str) -> str:
    """Extract and pretty-print an XML part from a docx."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        raw = zf.read(part_name)
    root = etree.fromstring(raw)
    return etree.tostring(
        root,
        pretty_print=True,
        encoding="unicode",
    )


def _generate(spec: DocumentSpec) -> Path:
    """Generate a docx in a temp dir and return its path."""
    tmpdir = tempfile.mkdtemp()
    path = Path(tmpdir) / "test.docx"
    DocumentGenerator(spec).generate(path)
    return path


def test_document_xml_simple_paragraph(snapshot):
    """Snapshot document.xml for a simple paragraph."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph("Hello world")
    path = _generate(spec)
    xml = _extract_xml(path, "word/document.xml")
    assert xml == snapshot


def test_comments_xml_with_reply(snapshot):
    """Snapshot comments.xml for a comment with reply."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Test paragraph with comment",
        comments=[
            Comment(
                text="Main comment",
                anchor_text="comment",
                replies=[
                    CommentReply(text="Reply text"),
                ],
            ),
        ],
    )
    path = _generate(spec)
    xml = _extract_xml(path, "word/comments.xml")
    assert xml == snapshot


def test_numbering_xml_legal_list(snapshot):
    """Snapshot numbering.xml for legal-list numbering."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Item 1",
        numbering=NumberedParagraph(level=0),
    )
    spec.add_paragraph(
        "Item 1.1",
        numbering=NumberedParagraph(level=1),
    )
    path = _generate(spec)
    xml = _extract_xml(path, "word/numbering.xml")
    assert xml == snapshot


def test_styles_xml_heading_numbering(snapshot):
    """Snapshot styles.xml for heading numbering."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph("Chapter One", heading_level=1)
    spec.add_paragraph("Section A", heading_level=2)
    path = _generate(spec)
    xml = _extract_xml(path, "word/styles.xml")
    assert xml == snapshot


def test_content_types_all_features(snapshot):
    """Snapshot [Content_Types].xml with all features."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Numbered item",
        numbering=NumberedParagraph(level=0),
    )
    spec.add_paragraph(
        "Commented text",
        comments=[
            Comment(
                text="A comment",
                anchor_text="Commented",
            )
        ],
    )
    spec.add_paragraph("Heading", heading_level=1)
    path = _generate(spec)
    xml = _extract_xml(path, "[Content_Types].xml")
    assert xml == snapshot
