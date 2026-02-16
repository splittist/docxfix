"""Typed specification models for docx fixtures."""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum


class ChangeType(Enum):
    """Type of tracked change."""

    INSERTION = "insertion"
    DELETION = "deletion"


@dataclass
class TrackedChange:
    """Specification for a tracked change (insertion or deletion).

    For insertions, ``insert_after`` identifies the substring of the paragraph
    text after which the inserted text is placed.  When empty, the insertion
    is appended after the paragraph text (or emitted standalone when the
    paragraph text is also empty).

    For deletions, ``text`` is located within the paragraph text and wrapped
    in a ``<w:del>`` element.  The surrounding text is emitted as plain runs.
    """

    change_type: ChangeType
    text: str
    author: str = "Test User"
    date: datetime | None = None
    revision_id: int = 1
    insert_after: str = ""

    def __post_init__(self) -> None:
        """Set default date if not provided."""
        if self.date is None:
            self.date = datetime.now()


@dataclass
class CommentReply:
    """Specification for a comment reply."""

    text: str
    author: str = "Test User"
    date: datetime | None = None

    def __post_init__(self) -> None:
        """Set default date if not provided."""
        if self.date is None:
            self.date = datetime.now()


@dataclass
class Comment:
    """Specification for a modern threaded comment."""

    text: str
    anchor_text: str
    author: str = "Test User"
    date: datetime | None = None
    replies: list[CommentReply] = field(default_factory=list)
    resolved: bool = False

    def __post_init__(self) -> None:
        """Set default date if not provided."""
        if self.date is None:
            self.date = datetime.now()


@dataclass
class NumberingLevel:
    """Specification for a numbering level."""

    level: int
    format: str = "%1."
    start: int = 1


@dataclass
class NumberedParagraph:
    """Specification for numbering properties of a paragraph."""

    level: int = 0
    numbering_id: int = 1


class PageOrientation(Enum):
    """Supported section page orientations."""

    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


@dataclass
class HeaderFooterSet:
    """Section header/footer text variants."""

    default: str | None = None
    first: str | None = None
    even: str | None = None


@dataclass
class SectionSpec:
    """Specification for a section starting at a paragraph index."""

    start_paragraph: int
    break_type: str = "nextPage"
    orientation: PageOrientation = PageOrientation.PORTRAIT
    restart_page_numbering: bool = False
    page_number_start: int | None = None
    headers: HeaderFooterSet = field(default_factory=HeaderFooterSet)
    footers: HeaderFooterSet = field(default_factory=HeaderFooterSet)

    def __post_init__(self) -> None:
        """Validate section settings."""
        if self.start_paragraph < 0:
            raise ValueError("Section start_paragraph must be >= 0")
        if self.page_number_start is not None and self.page_number_start < 1:
            raise ValueError("page_number_start must be >= 1")


@dataclass
class Paragraph:
    """Specification for a paragraph in the document."""

    text: str
    tracked_changes: list[TrackedChange] = field(default_factory=list)
    comments: list[Comment] = field(default_factory=list)
    numbering: NumberedParagraph | None = None
    heading_level: int | None = None  # 1-4 → Heading1-Heading4


@dataclass
class DocumentSpec:
    """Top-level specification for a docx fixture."""

    paragraphs: list[Paragraph] = field(default_factory=list)
    title: str = "Test Document"
    author: str = "Test User"
    seed: int | None = None
    sections: list[SectionSpec] = field(default_factory=list)

    def __post_init__(self) -> None:
        """Ensure section list always includes an initial section."""
        if not self.sections:
            self.sections = [SectionSpec(start_paragraph=0)]
        elif not any(section.start_paragraph == 0 for section in self.sections):
            self.sections.insert(0, SectionSpec(start_paragraph=0))

    def add_paragraph(
        self,
        text: str,
        tracked_changes: list[TrackedChange] | None = None,
        comments: list[Comment] | None = None,
        numbering: NumberedParagraph | None = None,
        heading_level: int | None = None,
    ) -> "DocumentSpec":
        """
        Add a paragraph to the document.

        Args:
            text: The paragraph text
            tracked_changes: Optional list of tracked changes
            comments: Optional list of comments
            numbering: Optional numbering configuration
            heading_level: Optional heading level (1-4) for styled numbering

        Returns:
            Self for method chaining
        """
        self.paragraphs.append(
            Paragraph(
                text=text,
                tracked_changes=tracked_changes or [],
                comments=comments or [],
                numbering=numbering,
                heading_level=heading_level,
            )
        )
        return self

    def add_section(
        self,
        start_paragraph: int,
        break_type: str = "nextPage",
        orientation: PageOrientation = PageOrientation.PORTRAIT,
        restart_page_numbering: bool = False,
        page_number_start: int | None = None,
        headers: HeaderFooterSet | None = None,
        footers: HeaderFooterSet | None = None,
    ) -> "DocumentSpec":
        """Add a section definition to the document."""
        self.sections.append(
            SectionSpec(
                start_paragraph=start_paragraph,
                break_type=break_type,
                orientation=orientation,
                restart_page_numbering=restart_page_numbering,
                page_number_start=page_number_start,
                headers=headers or HeaderFooterSet(),
                footers=footers or HeaderFooterSet(),
            )
        )
        return self
