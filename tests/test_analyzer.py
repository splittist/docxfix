"""Tests for docxfix.analyzer module."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pytest
from typer.testing import CliRunner

from docxfix.analyzer import AnalysisResult, CommentSummary, TrackedChangeSummary, analyze_docx
from docxfix.cli import app
from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    NumberedParagraph,
    Paragraph,
    SectionSpec,
    TrackedChange,
)


def _generate(spec: DocumentSpec) -> Path:
    """Generate a .docx into a temp file and return its path."""
    tmpdir = tempfile.mkdtemp()
    out = Path(tmpdir) / "test.docx"
    DocumentGenerator(spec).generate(str(out))
    return out


# ---------------------------------------------------------------------------
# Basic structure
# ---------------------------------------------------------------------------

def test_analyze_empty_doc():
    spec = DocumentSpec(title="Empty", author="A")
    path = _generate(spec)
    result = analyze_docx(path)
    assert isinstance(result, AnalysisResult)
    assert result.paragraph_count >= 0


def test_analyze_paragraph_count():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("First paragraph")
    spec.add_paragraph("Second paragraph")
    spec.add_paragraph("Third paragraph")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.paragraph_count >= 3


def test_analyze_file_not_found():
    with pytest.raises(FileNotFoundError):
        analyze_docx("/nonexistent/path/file.docx")


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------

def test_analyze_heading_counts():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Heading 1 text", heading_level=1)
    spec.add_paragraph("Heading 2 text", heading_level=2)
    spec.add_paragraph("Heading 2 again", heading_level=2)
    spec.add_paragraph("Body text")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.heading_counts.get("Heading1", 0) == 1
    assert result.heading_counts.get("Heading2", 0) == 2


def test_analyze_no_headings():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Just body text")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.heading_counts == {}


# ---------------------------------------------------------------------------
# Tracked changes
# ---------------------------------------------------------------------------

def test_analyze_tracked_changes_counts():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Some text with changes",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                author="Alice",
                text="inserted text",
            ),
            TrackedChange(
                change_type=ChangeType.DELETION,
                author="Bob",
                text="deleted text",
            ),
        ],
    )
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.tracked_changes.insertion_count >= 1
    assert result.tracked_changes.deletion_count >= 1


def test_analyze_tracked_changes_authors():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Text with changes",
        tracked_changes=[
            TrackedChange(change_type=ChangeType.INSERTION, author="Alice", text="hi"),
            TrackedChange(change_type=ChangeType.DELETION, author="Bob", text="bye"),
        ],
    )
    path = _generate(spec)
    result = analyze_docx(path)
    assert "Alice" in result.tracked_changes.authors
    assert "Bob" in result.tracked_changes.authors
    assert result.tracked_changes.authors == sorted(result.tracked_changes.authors)


def test_analyze_no_tracked_changes():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Clean text")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.tracked_changes.insertion_count == 0
    assert result.tracked_changes.deletion_count == 0
    assert result.tracked_changes.authors == []


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------

def test_analyze_comments_basic():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Some comment text",
        comments=[Comment(author="Alice", text="A comment", anchor_text="comment")],
    )
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.comments.total_count >= 1
    assert result.comments.thread_count >= 1
    assert result.comments.reply_count == 0
    assert "Alice" in result.comments.authors


def test_analyze_comments_with_replies():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Thread text here",
        comments=[
            Comment(
                author="Alice",
                text="Main comment",
                anchor_text="Thread",
                replies=[
                    CommentReply(author="Bob", text="Reply 1"),
                    CommentReply(author="Carol", text="Reply 2"),
                ],
            )
        ],
    )
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.comments.thread_count == 1
    assert result.comments.reply_count == 2
    assert result.comments.total_count == 3
    assert "Alice" in result.comments.authors
    assert "Bob" in result.comments.authors


def test_analyze_multiple_comment_threads():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "First second third",
        comments=[
            Comment(author="Alice", text="C1", anchor_text="First"),
            Comment(author="Bob", text="C2", anchor_text="second"),
        ],
    )
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.comments.thread_count == 2
    assert result.comments.total_count == 2


def test_analyze_no_comments():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Plain text")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.comments.total_count == 0
    assert result.comments.thread_count == 0
    assert result.comments.reply_count == 0
    assert result.comments.authors == []


# ---------------------------------------------------------------------------
# Numbering
# ---------------------------------------------------------------------------

def test_analyze_numbered_paragraphs():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Item one", numbering=NumberedParagraph(level=0))
    spec.add_paragraph("Item two", numbering=NumberedParagraph(level=0))
    spec.add_paragraph("Sub item", numbering=NumberedParagraph(level=1))
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.numbered_paragraph_count == 3


def test_analyze_no_numbering():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Plain")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.numbered_paragraph_count == 0


# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------

def test_analyze_single_section():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("One section only")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.section_count == 1


def test_analyze_multiple_sections():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Section 1 text")
    spec.add_paragraph("Section 2 text")
    spec.add_section(start_paragraph=1, break_type="nextPage")
    path = _generate(spec)
    result = analyze_docx(path)
    assert result.section_count >= 2


def spec_para(text: str) -> Paragraph:
    """Helper to create a Paragraph."""
    return Paragraph(text=text)


# ---------------------------------------------------------------------------
# to_dict / to_json
# ---------------------------------------------------------------------------

def test_analysis_result_to_dict():
    result = AnalysisResult(paragraph_count=5, section_count=1)
    d = result.to_dict()
    assert d["paragraph_count"] == 5
    assert d["section_count"] == 1
    assert "tracked_changes" in d
    assert "comments" in d


def test_analysis_result_to_json():
    result = AnalysisResult(paragraph_count=3)
    j = result.to_json()
    parsed = json.loads(j)
    assert parsed["paragraph_count"] == 3


# ---------------------------------------------------------------------------
# CLI integration
# ---------------------------------------------------------------------------

runner = CliRunner()


def test_cli_analyze_text_output():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Hello world")
    path = _generate(spec)
    result = runner.invoke(app, ["analyze", str(path)])
    assert result.exit_code == 0
    assert "Paragraphs" in result.output
    assert "Tracked changes" in result.output
    assert "Comments" in result.output
    assert "Sections" in result.output


def test_cli_analyze_json_output():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph("Hello world")
    path = _generate(spec)
    result = runner.invoke(app, ["analyze", str(path), "--output-format", "json"])
    assert result.exit_code == 0
    parsed = json.loads(result.output.strip())
    assert "paragraph_count" in parsed
    assert "tracked_changes" in parsed
    assert "comments" in parsed


def test_cli_analyze_missing_file():
    result = runner.invoke(app, ["analyze", "/nonexistent/path/file.docx"])
    assert result.exit_code != 0


def test_cli_analyze_shows_tracked_change_details():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Text with insertion",
        tracked_changes=[
            TrackedChange(change_type=ChangeType.INSERTION, author="Alice", text="hi"),
        ],
    )
    path = _generate(spec)
    result = runner.invoke(app, ["analyze", str(path)])
    assert result.exit_code == 0
    assert "Alice" in result.output


def test_cli_analyze_shows_comment_authors():
    spec = DocumentSpec(title="T", author="A")
    spec.add_paragraph(
        "Comment text",
        comments=[Comment(author="Reviewer", text="Note", anchor_text="Comment")],
    )
    path = _generate(spec)
    result = runner.invoke(app, ["analyze", str(path)])
    assert result.exit_code == 0
    assert "Reviewer" in result.output
