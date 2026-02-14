"""Tests for the spec module."""

from datetime import datetime

import pytest

from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    NumberedParagraph,
    Paragraph,
    TrackedChange,
)


def test_tracked_change_insertion():
    """Test creating an insertion tracked change."""
    change = TrackedChange(
        change_type=ChangeType.INSERTION,
        text="inserted text",
        author="John Doe",
    )
    assert change.change_type == ChangeType.INSERTION
    assert change.text == "inserted text"
    assert change.author == "John Doe"
    assert isinstance(change.date, datetime)


def test_tracked_change_deletion():
    """Test creating a deletion tracked change."""
    change = TrackedChange(
        change_type=ChangeType.DELETION,
        text="deleted text",
        author="Jane Smith",
    )
    assert change.change_type == ChangeType.DELETION
    assert change.text == "deleted text"
    assert change.author == "Jane Smith"


def test_tracked_change_with_custom_date():
    """Test tracked change with custom date."""
    custom_date = datetime(2024, 1, 1, 12, 0, 0)
    change = TrackedChange(
        change_type=ChangeType.INSERTION,
        text="test",
        date=custom_date,
    )
    assert change.date == custom_date


def test_comment_reply():
    """Test creating a comment reply."""
    reply = CommentReply(text="This is a reply", author="Reviewer")
    assert reply.text == "This is a reply"
    assert reply.author == "Reviewer"
    assert isinstance(reply.date, datetime)


def test_comment_with_replies():
    """Test creating a comment with replies."""
    comment = Comment(
        text="Main comment",
        anchor_text="highlighted text",
        author="Author",
        replies=[
            CommentReply(text="First reply", author="Reviewer 1"),
            CommentReply(text="Second reply", author="Reviewer 2"),
        ],
    )
    assert comment.text == "Main comment"
    assert comment.anchor_text == "highlighted text"
    assert len(comment.replies) == 2
    assert not comment.resolved


def test_comment_resolved():
    """Test creating a resolved comment."""
    comment = Comment(
        text="Resolved issue",
        anchor_text="text",
        resolved=True,
    )
    assert comment.resolved


def test_numbered_paragraph():
    """Test creating a numbered paragraph."""
    numbered = NumberedParagraph(
        text="Numbered item",
        level=1,
        numbering_id=2,
    )
    assert numbered.text == "Numbered item"
    assert numbered.level == 1
    assert numbered.numbering_id == 2


def test_paragraph_simple():
    """Test creating a simple paragraph."""
    para = Paragraph(text="Simple paragraph")
    assert para.text == "Simple paragraph"
    assert para.tracked_changes == []
    assert para.comments == []
    assert para.numbering is None


def test_paragraph_with_tracked_changes():
    """Test creating a paragraph with tracked changes."""
    change = TrackedChange(
        change_type=ChangeType.INSERTION,
        text="new text",
    )
    para = Paragraph(
        text="Main text",
        tracked_changes=[change],
    )
    assert len(para.tracked_changes) == 1
    assert para.tracked_changes[0].text == "new text"


def test_document_spec_empty():
    """Test creating an empty document spec."""
    spec = DocumentSpec()
    assert spec.paragraphs == []
    assert spec.title == "Test Document"
    assert spec.author == "Test User"
    assert spec.seed is None


def test_document_spec_add_paragraph():
    """Test adding paragraphs to document spec."""
    spec = DocumentSpec()
    spec.add_paragraph("First paragraph")
    spec.add_paragraph("Second paragraph")

    assert len(spec.paragraphs) == 2
    assert spec.paragraphs[0].text == "First paragraph"
    assert spec.paragraphs[1].text == "Second paragraph"


def test_document_spec_add_paragraph_with_changes():
    """Test adding paragraph with tracked changes."""
    spec = DocumentSpec()
    change = TrackedChange(
        change_type=ChangeType.DELETION,
        text="removed",
    )
    spec.add_paragraph(
        "Text with changes",
        tracked_changes=[change],
    )

    assert len(spec.paragraphs) == 1
    assert len(spec.paragraphs[0].tracked_changes) == 1


def test_document_spec_method_chaining():
    """Test method chaining with add_paragraph."""
    spec = (
        DocumentSpec()
        .add_paragraph("First")
        .add_paragraph("Second")
        .add_paragraph("Third")
    )

    assert len(spec.paragraphs) == 3


def test_document_spec_with_seed():
    """Test creating document spec with seed."""
    spec = DocumentSpec(seed=42)
    assert spec.seed == 42
