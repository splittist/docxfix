"""Integration tests for tracked change generation against golden corpus.

Validates that generated documents match the XML structure observed in:
- corpus/single-insertion.docx
- corpus/single-deletion.docx
- corpus/mixed-insert-delete.docx
"""

import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import ChangeType, DocumentSpec, TrackedChange
from docxfix.validator import validate_docx

CORPUS_DIR = Path(__file__).resolve().parent.parent / "corpus"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _generate_and_parse(spec: DocumentSpec):
    """Generate a .docx from spec and return parsed document.xml root."""
    tmpdir_ctx = tempfile.TemporaryDirectory()
    tmpdir = tmpdir_ctx.__enter__()
    output_path = Path(tmpdir) / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    validate_docx(output_path)
    docx_zip = zipfile.ZipFile(output_path, "r")
    root = etree.fromstring(docx_zip.read("word/document.xml"))
    return root, docx_zip, tmpdir_ctx


# ---------------------------------------------------------------------------
# Golden corpus structure tests (read the corpus files directly)
# ---------------------------------------------------------------------------


def test_corpus_single_insertion_structure():
    """Verify the golden single-insertion.docx has the expected XML shape."""
    with zipfile.ZipFile(CORPUS_DIR / "single-insertion.docx") as z:
        root = etree.fromstring(z.read("word/document.xml"))

    body = root.find("w:body", NS)
    paras = body.findall("w:p", NS)
    assert len(paras) >= 1

    # First paragraph should have exactly one w:ins
    p = paras[0]
    insertions = p.findall("w:ins", NS)
    assert len(insertions) == 1

    ins = insertions[0]
    assert ins.get(f"{{{W_NS}}}author") is not None
    assert ins.get(f"{{{W_NS}}}id") is not None

    # Insertion contains w:t (not w:delText)
    texts = ins.findall(".//w:t", NS)
    assert len(texts) >= 1
    assert "single insertion" in texts[0].text

    # No deletions in this document
    assert len(p.findall("w:del", NS)) == 0


def test_corpus_single_deletion_structure():
    """Verify the golden single-deletion.docx has the expected XML shape."""
    with zipfile.ZipFile(CORPUS_DIR / "single-deletion.docx") as z:
        root = etree.fromstring(z.read("word/document.xml"))

    body = root.find("w:body", NS)
    p = body.findall("w:p", NS)[0]

    deletions = p.findall("w:del", NS)
    assert len(deletions) == 1

    del_elem = deletions[0]
    assert del_elem.get(f"{{{W_NS}}}author") is not None

    # Deletion contains w:delText (not w:t)
    del_texts = del_elem.findall(".//w:delText", NS)
    assert len(del_texts) >= 1
    assert "dolor sit amet" in del_texts[0].text

    # No insertions
    assert len(p.findall("w:ins", NS)) == 0


def test_corpus_mixed_insert_delete_structure():
    """Verify the golden mixed-insert-delete.docx has both ins and del."""
    with zipfile.ZipFile(CORPUS_DIR / "mixed-insert-delete.docx") as z:
        root = etree.fromstring(z.read("word/document.xml"))

    body = root.find("w:body", NS)
    p = body.findall("w:p", NS)[0]

    insertions = p.findall("w:ins", NS)
    deletions = p.findall("w:del", NS)
    assert len(insertions) == 1
    assert len(deletions) == 1

    # IDs should be unique
    ins_id = insertions[0].get(f"{{{W_NS}}}id")
    del_id = deletions[0].get(f"{{{W_NS}}}id")
    assert ins_id != del_id


# ---------------------------------------------------------------------------
# Generator tests matching golden corpus patterns
# ---------------------------------------------------------------------------


def test_generate_single_insertion_matches_corpus():
    """Generate a document matching single-insertion.docx pattern.

    Corpus: plain run "Lorem ipsum dolor " → <w:ins>"single insertion " → plain run "sit amet..."
    """
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="single insertion ",
                author="Author",
                date=datetime(2026, 2, 15, 12, 0, 0),
                insert_after="Lorem ipsum dolor ",
            )
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        # Should have exactly one insertion
        insertions = p.findall("w:ins", NS)
        assert len(insertions) == 1

        ins = insertions[0]
        assert ins.get(f"{{{W_NS}}}author") == "Author"

        # Insertion text
        ins_texts = ins.findall(".//w:t", NS)
        assert len(ins_texts) == 1
        assert ins_texts[0].text == "single insertion "
        # Whitespace preservation
        assert (
            ins_texts[0].get("{http://www.w3.org/XML/1998/namespace}space")
            == "preserve"
        )

        # Should have plain text runs before and after the insertion
        # Get all direct children of paragraph that are runs or tracked changes
        children = [
            c
            for c in p
            if c.tag in (f"{{{W_NS}}}r", f"{{{W_NS}}}ins", f"{{{W_NS}}}del")
        ]
        # Expect: plain run, insertion, plain run
        assert len(children) == 3
        assert children[0].tag == f"{{{W_NS}}}r"
        assert children[1].tag == f"{{{W_NS}}}ins"
        assert children[2].tag == f"{{{W_NS}}}r"

        # Check before-text
        before_text = children[0].find("w:t", NS).text
        assert before_text == "Lorem ipsum dolor "

        # Check after-text
        after_text = children[2].find("w:t", NS).text
        assert after_text == "sit amet, consectetuer adipiscing elit."

        # No deletions
        assert len(p.findall("w:del", NS)) == 0
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_generate_single_deletion_matches_corpus():
    """Generate a document matching single-deletion.docx pattern.

    Corpus: plain run "Lorem ipsum" → <w:del>" dolor sit amet" → plain run ", consectetuer..."
    """
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" dolor sit amet",
                author="Author",
                date=datetime(2026, 2, 15, 12, 0, 0),
            )
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        # Should have exactly one deletion
        deletions = p.findall("w:del", NS)
        assert len(deletions) == 1

        del_elem = deletions[0]
        assert del_elem.get(f"{{{W_NS}}}author") == "Author"

        # Deletion uses w:delText
        del_texts = del_elem.findall(".//w:delText", NS)
        assert len(del_texts) == 1
        assert del_texts[0].text == " dolor sit amet"
        # Whitespace preservation on delText
        assert (
            del_texts[0].get("{http://www.w3.org/XML/1998/namespace}space")
            == "preserve"
        )

        # Check paragraph structure: plain run, deletion, plain run
        children = [
            c
            for c in p
            if c.tag in (f"{{{W_NS}}}r", f"{{{W_NS}}}ins", f"{{{W_NS}}}del")
        ]
        assert len(children) == 3
        assert children[0].tag == f"{{{W_NS}}}r"
        assert children[1].tag == f"{{{W_NS}}}del"
        assert children[2].tag == f"{{{W_NS}}}r"

        # Before text
        assert children[0].find("w:t", NS).text == "Lorem ipsum"
        # After text
        assert (
            children[2].find("w:t", NS).text
            == ", consectetuer adipiscing elit."
        )

        # No insertions
        assert len(p.findall("w:ins", NS)) == 0
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_generate_mixed_insert_delete_matches_corpus():
    """Generate a document matching mixed-insert-delete.docx pattern.

    Corpus: "Lorem " → <w:ins>"dolor sit amet " → "ipsum" → <w:del>" dolor sit amet" → ", consectetuer..."
    """
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="dolor sit amet ",
                author="Author",
                date=datetime(2026, 2, 15, 12, 0, 0),
                insert_after="Lorem ",
            ),
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" dolor sit amet",
                author="Author",
                date=datetime(2026, 2, 15, 12, 0, 0),
            ),
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        insertions = p.findall("w:ins", NS)
        deletions = p.findall("w:del", NS)
        assert len(insertions) == 1
        assert len(deletions) == 1

        # IDs should be unique
        ins_id = insertions[0].get(f"{{{W_NS}}}id")
        del_id = deletions[0].get(f"{{{W_NS}}}id")
        assert ins_id != del_id

        # Check full structure: "Lorem " → ins → "ipsum" → del → ", consectetuer..."
        children = [
            c
            for c in p
            if c.tag in (f"{{{W_NS}}}r", f"{{{W_NS}}}ins", f"{{{W_NS}}}del")
        ]
        assert len(children) == 5

        assert children[0].tag == f"{{{W_NS}}}r"  # "Lorem "
        assert children[1].tag == f"{{{W_NS}}}ins"  # "dolor sit amet "
        assert children[2].tag == f"{{{W_NS}}}r"  # "ipsum"
        assert children[3].tag == f"{{{W_NS}}}del"  # " dolor sit amet"
        assert children[4].tag == f"{{{W_NS}}}r"  # ", consectetuer..."

        # Verify text content
        assert children[0].find("w:t", NS).text == "Lorem "
        assert insertions[0].findall(".//w:t", NS)[0].text == "dolor sit amet "
        assert children[2].find("w:t", NS).text == "ipsum"
        assert deletions[0].findall(".//w:delText", NS)[0].text == " dolor sit amet"
        assert (
            children[4].find("w:t", NS).text
            == ", consectetuer adipiscing elit."
        )
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_generate_tracked_change_unique_ids():
    """Each tracked change element gets a unique w:id."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Some text here and more text.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="inserted ",
                author="A",
                date=datetime(2026, 1, 1),
                insert_after="Some ",
            ),
            TrackedChange(
                change_type=ChangeType.DELETION,
                text=" and more",
                author="B",
                date=datetime(2026, 1, 2),
            ),
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        all_ids = []
        for elem in p.findall("w:ins", NS) + p.findall("w:del", NS):
            all_ids.append(elem.get(f"{{{W_NS}}}id"))

        assert len(all_ids) == 2
        assert len(set(all_ids)) == 2  # all unique
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_generate_tracked_change_validates():
    """Generated tracked change documents pass structural validation."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Hello world, this is a test.",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="beautiful ",
                author="Author",
                date=datetime(2026, 2, 15, 12, 0, 0),
                insert_after="Hello ",
            ),
        ],
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        # Should not raise
        validate_docx(output_path)


def test_generate_legacy_tracked_changes_without_base_text():
    """Tracked changes without base text still work (legacy behaviour)."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="inserted",
                author="A",
                date=datetime(2026, 1, 1),
            ),
            TrackedChange(
                change_type=ChangeType.DELETION,
                text="deleted",
                author="B",
                date=datetime(2026, 1, 2),
            ),
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        assert len(p.findall("w:ins", NS)) == 1
        assert len(p.findall("w:del", NS)) == 1
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)


def test_xml_space_preserve_on_whitespace_text():
    """Text elements with leading/trailing whitespace get xml:space='preserve'."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph(
        "Before after",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text=" middle ",
                author="A",
                date=datetime(2026, 1, 1),
                insert_after="Before",
            ),
        ],
    )

    root, docx_zip, ctx = _generate_and_parse(spec)
    try:
        body = root.find("w:body", NS)
        p = body.findall("w:p", NS)[0]

        ins = p.findall("w:ins", NS)[0]
        t = ins.findall(".//w:t", NS)[0]
        assert t.text == " middle "
        assert (
            t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
        )

        # "Before" has no trailing space, so no xml:space
        children = [c for c in p if c.tag == f"{{{W_NS}}}r"]
        before_t = children[0].find("w:t", NS)
        assert before_t.text == "Before"
        assert (
            before_t.get("{http://www.w3.org/XML/1998/namespace}space") is None
        )
    finally:
        docx_zip.close()
        ctx.__exit__(None, None, None)
