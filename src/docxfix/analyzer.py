"""Analyze existing .docx files and return a summary of characteristics."""

from __future__ import annotations

import json
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass
class TrackedChangeSummary:
    insertion_count: int = 0
    deletion_count: int = 0
    authors: list[str] = field(default_factory=list)


@dataclass
class CommentSummary:
    thread_count: int = 0
    reply_count: int = 0
    total_count: int = 0
    authors: list[str] = field(default_factory=list)


@dataclass
class AnalysisResult:
    paragraph_count: int = 0
    heading_counts: dict[str, int] = field(default_factory=dict)
    tracked_changes: TrackedChangeSummary = field(
        default_factory=TrackedChangeSummary
    )
    comments: CommentSummary = field(default_factory=CommentSummary)
    numbered_paragraph_count: int = 0
    section_count: int = 0

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), indent=2)


def analyze_docx(path: str | Path) -> AnalysisResult:
    """Analyze a .docx file and return a summary of its characteristics."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    result = AnalysisResult()

    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        if "word/document.xml" in names:
            doc_tree = etree.fromstring(z.read("word/document.xml"))
            _analyze_document(doc_tree, result)

        if "word/comments.xml" in names:
            comments_tree = etree.fromstring(z.read("word/comments.xml"))
            _analyze_comments_xml(comments_tree, result)

        if "word/commentsExtended.xml" in names:
            ext_tree = etree.fromstring(z.read("word/commentsExtended.xml"))
            _analyze_comments_extended(ext_tree, result)

    return result


def _analyze_document(doc_tree: etree._Element, result: AnalysisResult) -> None:
    """Extract paragraph, tracked change, and section stats from document.xml."""
    paragraphs = doc_tree.findall(f".//{{{_W}}}p")
    result.paragraph_count = len(paragraphs)

    heading_counts: dict[str, int] = {}
    numbered_count = 0
    for para in paragraphs:
        p_pr = para.find(f"{{{_W}}}pPr")
        if p_pr is not None:
            p_style = p_pr.find(f"{{{_W}}}pStyle")
            if p_style is not None:
                style_val = p_style.get(f"{{{_W}}}val", "")
                if style_val.startswith("Heading"):
                    heading_counts[style_val] = heading_counts.get(style_val, 0) + 1
            if p_pr.find(f"{{{_W}}}numPr") is not None:
                numbered_count += 1

    result.heading_counts = heading_counts
    result.numbered_paragraph_count = numbered_count

    # Tracked changes
    insertions = doc_tree.findall(f".//{{{_W}}}ins")
    deletions = doc_tree.findall(f".//{{{_W}}}del")
    result.tracked_changes.insertion_count = len(insertions)
    result.tracked_changes.deletion_count = len(deletions)
    tc_authors: set[str] = set()
    for elem in insertions + deletions:
        author = elem.get(f"{{{_W}}}author")
        if author:
            tc_authors.add(author)
    result.tracked_changes.authors = sorted(tc_authors)

    # Section count: 1 for each body-level sectPr plus sectPr inside pPr
    body = doc_tree.find(f"{{{_W}}}body")
    if body is not None:
        sect_count = 1 if body.find(f"{{{_W}}}sectPr") is not None else 0
        for para in paragraphs:
            p_pr = para.find(f"{{{_W}}}pPr")
            if p_pr is not None and p_pr.find(f"{{{_W}}}sectPr") is not None:
                sect_count += 1
        result.section_count = sect_count


def _analyze_comments_xml(
    comments_tree: etree._Element, result: AnalysisResult
) -> None:
    """Extract comment count and authors from comments.xml."""
    comments = comments_tree.findall(f"{{{_W}}}comment")
    result.comments.total_count = len(comments)
    authors: set[str] = set()
    for comment in comments:
        author = comment.get(f"{{{_W}}}author")
        if author:
            authors.add(author)
    result.comments.authors = sorted(authors)


def _analyze_comments_extended(
    ext_tree: etree._Element, result: AnalysisResult
) -> None:
    """Compute thread vs reply counts from commentsExtended.xml."""
    reply_count = sum(
        1
        for elem in ext_tree
        if elem.get(f"{{{_W15}}}paraIdParent") is not None
    )
    result.comments.reply_count = reply_count
    result.comments.thread_count = result.comments.total_count - reply_count
