"""Integration tests for combined fixture scenarios.

M5 integration hardening: validates that tracked changes, comments, and numbering
work correctly when combined in the same document.
"""

import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    NumberedParagraph,
    TrackedChange,
)
from docxfix.validator import validate_docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
NS = {"w": W_NS}
NS15 = {"w15": W15_NS}


def _generate_to_zip(spec: DocumentSpec):
    """Generate a .docx from spec, validate it, and return the ZipFile + tmpdir context."""
    tmpdir_ctx = tempfile.TemporaryDirectory()
    tmpdir = tmpdir_ctx.__enter__()
    output_path = Path(tmpdir) / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    validate_docx(output_path)
    docx_zip = zipfile.ZipFile(output_path, "r")
    return docx_zip, tmpdir_ctx


def _parse_part(docx_zip: zipfile.ZipFile, part_name: str) -> etree._Element:
    """Parse an XML part from a .docx zip."""
    return etree.fromstring(docx_zip.read(part_name))


# ---------------------------------------------------------------------------
# Combined: Tracked Changes + Comments
# ---------------------------------------------------------------------------


def test_tracked_changes_and_comments_in_same_paragraph():
    """A single paragraph with both a tracked insertion and a comment."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "The contract shall be governed by the laws of the State.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="exclusively ",
                author="Alice",
                date=datetime(2026, 3, 1, 10, 0, 0),
                insert_after="shall be ",
            ),
        ],
        comments=[
            Comment(
                text="Verify governing law clause with client.",
                anchor_text="governed by",
                author="Bob",
                date=datetime(2026, 3, 1, 11, 0, 0),
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        # Tracked insertion present
        insertions = p.findall("w:ins", NS)
        assert len(insertions) == 1
        ins_text = insertions[0].findall(".//w:t", NS)[0].text
        assert ins_text == "exclusively "

        # Comment anchoring present
        comment_starts = p.findall("w:commentRangeStart", NS)
        comment_ends = p.findall("w:commentRangeEnd", NS)
        assert len(comment_starts) >= 1
        assert len(comment_ends) >= 1

        # comments.xml exists and has the comment
        comments_root = _parse_part(docx_zip, "word/comments.xml")
        comments = comments_root.findall(".//w:comment", NS)
        assert len(comments) == 1
        assert comments[0].get(f"{{{W_NS}}}author") == "Bob"

        # commentsExtended.xml exists
        ext_root = _parse_part(docx_zip, "word/commentsExtended.xml")
        comment_exs = ext_root.findall(".//w15:commentEx", NS15)
        assert len(comment_exs) == 1
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_tracked_changes_and_comments_across_paragraphs():
    """Multiple paragraphs: one with tracked changes, another with comments."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "This paragraph has a deletion of some words.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" of some words",
                author="Alice",
                date=datetime(2026, 3, 1, 10, 0, 0),
            ),
        ],
    )
    spec.add_paragraph(
        "This paragraph has a comment on it.",
        comments=[
            Comment(
                text="Please review this paragraph.",
                anchor_text="comment on it",
                author="Bob",
                date=datetime(2026, 3, 1, 11, 0, 0),
                replies=[
                    CommentReply(
                        text="Reviewed and approved.",
                        author="Alice",
                        date=datetime(2026, 3, 1, 12, 0, 0),
                    ),
                ],
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        paras = body.findall("w:p", NS)
        assert len(paras) >= 2

        # First paragraph: deletion, no comments
        p1 = paras[0]
        assert len(p1.findall("w:del", NS)) == 1
        assert len(p1.findall("w:commentRangeStart", NS)) == 0

        # Second paragraph: comment, no tracked changes
        p2 = paras[1]
        assert len(p2.findall("w:ins", NS)) == 0
        assert len(p2.findall("w:del", NS)) == 0
        assert len(p2.findall("w:commentRangeStart", NS)) >= 1

        # comments.xml has 2 entries (main + reply)
        comments_root = _parse_part(docx_zip, "word/comments.xml")
        comments = comments_root.findall(".//w:comment", NS)
        assert len(comments) == 2
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


# ---------------------------------------------------------------------------
# Combined: Tracked Changes + Numbering
# ---------------------------------------------------------------------------


def test_tracked_changes_in_numbered_paragraph():
    """Numbered paragraphs can also contain tracked changes."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "First clause with an insertion here.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="important ",
                author="Lawyer A",
                date=datetime(2026, 3, 1, 10, 0, 0),
                insert_after="an ",
            ),
        ],
    )
    spec.add_paragraph(
        "Second clause is plain.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
    )
    spec.add_paragraph(
        "Sub-clause with a deletion of extra text.",
        numbering=NumberedParagraph(level=1, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" of extra text",
                author="Lawyer B",
                date=datetime(2026, 3, 1, 11, 0, 0),
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        paras = body.findall("w:p", NS)
        assert len(paras) >= 3

        # First paragraph: numbered + insertion
        p1 = paras[0]
        p1_ppr = p1.find("w:pPr", NS)
        assert p1_ppr is not None
        num_pr = p1_ppr.find("w:numPr", NS)
        assert num_pr is not None
        ilvl = num_pr.find("w:ilvl", NS)
        assert ilvl.get(f"{{{W_NS}}}val") == "0"
        assert len(p1.findall("w:ins", NS)) == 1

        # Third paragraph: numbered at level 1 + deletion
        p3 = paras[2]
        p3_ppr = p3.find("w:pPr", NS)
        num_pr_3 = p3_ppr.find("w:numPr", NS)
        ilvl_3 = num_pr_3.find("w:ilvl", NS)
        assert ilvl_3.get(f"{{{W_NS}}}val") == "1"
        assert len(p3.findall("w:del", NS)) == 1

        # numbering.xml and styles.xml present
        assert "word/numbering.xml" in docx_zip.namelist()
        assert "word/styles.xml" in docx_zip.namelist()
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


# ---------------------------------------------------------------------------
# Combined: Comments + Numbering
# ---------------------------------------------------------------------------


def test_comments_on_numbered_paragraphs():
    """Comments anchored on numbered paragraphs."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "First clause of the agreement.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        comments=[
            Comment(
                text="Needs client approval.",
                anchor_text="agreement",
                author="Reviewer",
                date=datetime(2026, 3, 1, 10, 0, 0),
                resolved=True,
            ),
        ],
    )
    spec.add_paragraph(
        "Definitions and interpretation.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
    )
    spec.add_paragraph(
        "The term 'Party' means the signatories.",
        numbering=NumberedParagraph(level=1, numbering_id=1),
        comments=[
            Comment(
                text="Add definition for Affiliate.",
                anchor_text="signatories",
                author="Reviewer",
                date=datetime(2026, 3, 1, 11, 0, 0),
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        paras = body.findall("w:p", NS)
        assert len(paras) >= 3

        # All three paragraphs have numbering
        for i, p in enumerate(paras[:3]):
            ppr = p.find("w:pPr", NS)
            assert ppr is not None, f"Paragraph {i} missing pPr"
            assert ppr.find("w:numPr", NS) is not None, f"Paragraph {i} missing numPr"

        # Comments present
        comments_root = _parse_part(docx_zip, "word/comments.xml")
        comments = comments_root.findall(".//w:comment", NS)
        assert len(comments) == 2

        # First comment is resolved
        ext_root = _parse_part(docx_zip, "word/commentsExtended.xml")
        comment_exs = ext_root.findall(".//w15:commentEx", NS15)
        assert len(comment_exs) == 2
        done_values = [ex.get(f"{{{W15_NS}}}done") for ex in comment_exs]
        assert "1" in done_values  # resolved comment
        assert "0" in done_values  # unresolved comment

        # Numbering and styles present
        assert "word/numbering.xml" in docx_zip.namelist()
        assert "word/styles.xml" in docx_zip.namelist()
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


# ---------------------------------------------------------------------------
# Combined: All Three Features (Tracked Changes + Comments + Numbering)
# ---------------------------------------------------------------------------


def test_all_features_combined():
    """A realistic legal document with tracked changes, comments, and numbering."""
    spec = DocumentSpec(seed=42)

    # Clause 1: plain numbered paragraph
    spec.add_paragraph(
        "This Agreement is entered into as of the Effective Date.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
    )

    # Clause 2: numbered paragraph with a tracked insertion and a comment
    spec.add_paragraph(
        "The Parties agree to the following terms and conditions.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="material ",
                author="Outside Counsel",
                date=datetime(2026, 3, 1, 14, 0, 0),
                insert_after="following ",
            ),
        ],
        comments=[
            Comment(
                text="Should we add 'material' here?",
                anchor_text="terms and conditions",
                author="Partner",
                date=datetime(2026, 3, 1, 15, 0, 0),
                replies=[
                    CommentReply(
                        text="Yes, it narrows the scope appropriately.",
                        author="Outside Counsel",
                        date=datetime(2026, 3, 1, 16, 0, 0),
                    ),
                ],
            ),
        ],
    )

    # Sub-clause 2.1: numbered sub-item with a deletion
    spec.add_paragraph(
        "Confidentiality obligations shall survive for a period of five years.",
        numbering=NumberedParagraph(level=1, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" for a period of",
                author="Outside Counsel",
                date=datetime(2026, 3, 1, 14, 30, 0),
            ),
        ],
    )

    # Sub-clause 2.2: nested item with a resolved comment
    spec.add_paragraph(
        "Non-compete restrictions apply within the Territory.",
        numbering=NumberedParagraph(level=1, numbering_id=1),
        comments=[
            Comment(
                text="Territory needs to be defined in definitions section.",
                anchor_text="Territory",
                author="Associate",
                date=datetime(2026, 3, 1, 17, 0, 0),
                resolved=True,
            ),
        ],
    )

    # Clause 3: plain numbered paragraph
    spec.add_paragraph(
        "This Agreement constitutes the entire understanding between the Parties.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        paras = body.findall("w:p", NS)
        assert len(paras) >= 5

        # --- Verify tracked changes ---
        # Paragraph 2: one insertion
        p2 = paras[1]
        assert len(p2.findall("w:ins", NS)) == 1
        ins_text = p2.findall("w:ins", NS)[0].findall(".//w:t", NS)[0].text
        assert ins_text == "material "

        # Paragraph 3: one deletion
        p3 = paras[2]
        assert len(p3.findall("w:del", NS)) == 1
        del_text = p3.findall("w:del", NS)[0].findall(".//w:delText", NS)[0].text
        assert del_text == " for a period of"

        # --- Verify comments ---
        comments_root = _parse_part(docx_zip, "word/comments.xml")
        comments = comments_root.findall(".//w:comment", NS)
        # 2 main comments + 1 reply = 3 total
        assert len(comments) == 3

        ext_root = _parse_part(docx_zip, "word/commentsExtended.xml")
        comment_exs = ext_root.findall(".//w15:commentEx", NS15)
        assert len(comment_exs) == 3

        # Check resolved state
        done_values = [ex.get(f"{{{W15_NS}}}done") for ex in comment_exs]
        assert done_values.count("1") == 1  # one resolved
        assert done_values.count("0") == 2  # two unresolved (main + reply)

        # Check reply threading
        parent_refs = [
            ex.get(f"{{{W15_NS}}}paraIdParent")
            for ex in comment_exs
            if f"{{{W15_NS}}}paraIdParent" in ex.attrib
        ]
        assert len(parent_refs) == 1  # only the reply has a parent

        # --- Verify numbering ---
        for i, p in enumerate(paras[:5]):
            ppr = p.find("w:pPr", NS)
            assert ppr is not None, f"Paragraph {i} missing pPr"
            num_pr = ppr.find("w:numPr", NS)
            assert num_pr is not None, f"Paragraph {i} missing numPr"

        # Check level assignments
        levels = []
        for p in paras[:5]:
            ilvl = p.find("w:pPr/w:numPr/w:ilvl", NS)
            levels.append(int(ilvl.get(f"{{{W_NS}}}val")))
        assert levels == [0, 0, 1, 1, 0]

        # --- Verify all parts present ---
        names = docx_zip.namelist()
        assert "word/comments.xml" in names
        assert "word/commentsExtended.xml" in names
        assert "word/numbering.xml" in names
        assert "word/styles.xml" in names

        # --- Verify numbering.xml structure ---
        num_root = _parse_part(docx_zip, "word/numbering.xml")
        abstract_nums = num_root.findall(".//w:abstractNum", NS)
        assert len(abstract_nums) >= 1
        nums = num_root.findall(".//w:num", NS)
        assert len(nums) >= 1
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_all_features_combined_validates():
    """Combined document passes structural validation without errors."""
    spec = DocumentSpec(seed=99)
    spec.add_paragraph(
        "Introduction paragraph.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="revised ",
                author="Editor",
                date=datetime(2026, 4, 1, 9, 0, 0),
                insert_after="Introduction ",
            ),
        ],
        comments=[
            Comment(
                text="Check formatting.",
                anchor_text="paragraph",
                author="Reviewer",
                date=datetime(2026, 4, 1, 10, 0, 0),
            ),
        ],
    )
    spec.add_paragraph(
        "Sub-item with deletion of redundant text here.",
        numbering=NumberedParagraph(level=1, numbering_id=1),
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" of redundant text",
                author="Editor",
                date=datetime(2026, 4, 1, 9, 30, 0),
            ),
        ],
    )
    spec.add_paragraph(
        "Plain paragraph without any features.",
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "combined.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        # Should not raise
        validate_docx(output_path)


def test_content_types_include_all_parts():
    """Content types XML includes entries for all feature parts."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Numbered paragraph with comment.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        comments=[
            Comment(
                text="A note.",
                anchor_text="comment",
                author="User",
                date=datetime(2026, 3, 1, 10, 0, 0),
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        ct_root = _parse_part(docx_zip, "[Content_Types].xml")
        overrides = ct_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
        )
        part_names = [o.get("PartName") for o in overrides]

        assert "/word/document.xml" in part_names
        assert "/word/comments.xml" in part_names
        assert "/word/commentsExtended.xml" in part_names
        assert "/word/commentsIds.xml" in part_names
        assert "/word/numbering.xml" in part_names
        assert "/word/styles.xml" in part_names
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_relationships_include_all_parts():
    """document.xml.rels includes relationships for all feature parts."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Numbered with comment and tracked change.",
        numbering=NumberedParagraph(level=0, numbering_id=1),
        comments=[
            Comment(
                text="Review needed.",
                anchor_text="comment",
                author="User",
                date=datetime(2026, 3, 1, 10, 0, 0),
            ),
        ],
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="also ",
                author="User",
                date=datetime(2026, 3, 1, 11, 0, 0),
                insert_after="Numbered ",
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        rels_root = _parse_part(docx_zip, "word/_rels/document.xml.rels")
        rels = rels_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        targets = [r.get("Target") for r in rels]

        assert "comments.xml" in targets
        assert "commentsExtended.xml" in targets
        assert "commentsIds.xml" in targets
        assert "numbering.xml" in targets
        assert "styles.xml" in targets
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_unique_ids_across_features():
    """Tracked change IDs and comment IDs don't collide in combined documents."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Text with insertion and comment.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="new ",
                author="A",
                date=datetime(2026, 3, 1, 10, 0, 0),
                insert_after="with ",
            ),
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" and",
                author="B",
                date=datetime(2026, 3, 1, 11, 0, 0),
            ),
        ],
        comments=[
            Comment(
                text="Note about insertion.",
                anchor_text="insertion",
                author="C",
                date=datetime(2026, 3, 1, 12, 0, 0),
                replies=[
                    CommentReply(
                        text="Acknowledged.",
                        author="D",
                        date=datetime(2026, 3, 1, 13, 0, 0),
                    ),
                ],
            ),
        ],
    )

    docx_zip, ctx = _generate_to_zip(spec)
    try:
        doc = _parse_part(docx_zip, "word/document.xml")
        body = doc.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        # Collect all w:id values from tracked changes
        tc_ids = set()
        for elem in p.findall("w:ins", NS) + p.findall("w:del", NS):
            tc_ids.add(elem.get(f"{{{W_NS}}}id"))

        # Collect comment IDs
        comment_ids = set()
        for elem in p.findall(".//w:commentRangeStart", NS):
            comment_ids.add(elem.get(f"{{{W_NS}}}id"))

        # All tracked change IDs are unique among themselves
        assert len(tc_ids) == 2

        # All comment IDs are unique among themselves
        assert len(comment_ids) >= 1
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)
