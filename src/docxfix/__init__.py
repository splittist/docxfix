"""
docxfix - A CLI utility for creating docx fixtures.

Creates docx fixtures with desirable characteristics for testing.
"""

__version__ = "0.1.0"

from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    HeaderFooterSet,
    NumberedParagraph,
    NumberingLevel,
    PageOrientation,
    Paragraph,
    SectionSpec,
    TrackedChange,
)
from docxfix.validator import DocumentValidator, ValidationError, validate_docx

__all__ = [
    "ChangeType",
    "Comment",
    "CommentReply",
    "DocumentGenerator",
    "DocumentSpec",
    "DocumentValidator",
    "HeaderFooterSet",
    "NumberedParagraph",
    "NumberingLevel",
    "PageOrientation",
    "Paragraph",
    "SectionSpec",
    "TrackedChange",
    "ValidationError",
    "validate_docx",
]
