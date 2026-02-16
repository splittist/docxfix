"""Tests for heading-based (styled) numbering in DocumentGenerator."""

import zipfile

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    Comment,
    DocumentSpec,
    NumberedParagraph,
    TrackedChange,
    ChangeType,
)
from docxfix.validator import validate_docx

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _generate(spec, tmp_path):
    """Generate a docx and return the ZipFile."""
    out = tmp_path / "test.docx"
    DocumentGenerator(spec).generate(out)
    return zipfile.ZipFile(out)


# ------------------------------------------------------------------
# Paragraph emission: pStyle only, no numPr in document.xml
# ------------------------------------------------------------------


def test_heading_paragraph_has_pstyle_no_numpr(tmp_path):
    """Heading paragraphs get pStyle=HeadingN with no numPr."""
    spec = DocumentSpec()
    spec.add_paragraph("Chapter One", heading_level=1)
    spec.add_paragraph("Section A", heading_level=2)

    with _generate(spec, tmp_path) as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        paras = doc.findall(".//w:p", NS)
        assert len(paras) == 2

        for para, expected_style in zip(paras, ["Heading1", "Heading2"]):
            pPr = para.find("w:pPr", NS)
            assert pPr is not None
            pStyle = pPr.find("w:pStyle", NS)
            assert pStyle is not None
            assert pStyle.get(f"{{{NS['w']}}}val") == expected_style
            # No explicit numPr in the paragraph
            assert pPr.find("w:numPr", NS) is None


def test_heading_all_four_levels(tmp_path):
    """Levels 1-4 all produce the correct HeadingN style."""
    spec = DocumentSpec()
    for lvl in range(1, 5):
        spec.add_paragraph(f"Level {lvl}", heading_level=lvl)

    with _generate(spec, tmp_path) as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        paras = doc.findall(".//w:p", NS)
        assert len(paras) == 4
        for i, para in enumerate(paras, start=1):
            style_val = para.find("w:pPr/w:pStyle", NS).get(f"{{{NS['w']}}}val")
            assert style_val == f"Heading{i}"


# ------------------------------------------------------------------
# styles.xml: Heading styles contain numPr
# ------------------------------------------------------------------


def test_styles_contain_heading_definitions(tmp_path):
    """styles.xml has Heading1-4 with embedded numPr."""
    spec = DocumentSpec()
    spec.add_paragraph("Title", heading_level=1)

    with _generate(spec, tmp_path) as z:
        styles = etree.fromstring(z.read("word/styles.xml"))

        for i in range(1, 5):
            style_id = f"Heading{i}"
            style = styles.find(
                f".//w:style[@w:styleId='{style_id}']",
                NS,
            )
            assert style is not None, f"Missing {style_id} in styles.xml"

            # basedOn Normal
            based = style.find("w:basedOn", NS)
            assert based is not None
            assert based.get(f"{{{NS['w']}}}val") == "Normal"

            # numPr inside pPr
            numPr = style.find("w:pPr/w:numPr", NS)
            assert numPr is not None, f"{style_id} missing numPr"

            numId = numPr.find("w:numId", NS)
            assert numId is not None
            # heading numId is "1" when no list numbering
            assert numId.get(f"{{{NS['w']}}}val") == "1"

            # ilvl: omitted for level 0, present for 1-3
            ilvl = numPr.find("w:ilvl", NS)
            if i == 1:
                assert ilvl is None, "Heading1 should omit ilvl (implicit 0)"
            else:
                assert ilvl is not None
                assert ilvl.get(f"{{{NS['w']}}}val") == str(i - 1)


def test_heading_styles_have_formatting(tmp_path):
    """Heading styles include font and bold properties."""
    spec = DocumentSpec()
    spec.add_paragraph("Title", heading_level=1)

    with _generate(spec, tmp_path) as z:
        styles = etree.fromstring(z.read("word/styles.xml"))
        h1 = styles.find(".//w:style[@w:styleId='Heading1']", NS)
        rPr = h1.find("w:rPr", NS)
        assert rPr is not None
        assert rPr.find("w:b", NS) is not None  # bold
        sz = rPr.find("w:sz", NS)
        assert sz is not None
        assert sz.get(f"{{{NS['w']}}}val") == "32"


# ------------------------------------------------------------------
# numbering.xml: pStyle back-references + correct numFmt/lvlText
# ------------------------------------------------------------------


def test_numbering_has_heading_abstract_num(tmp_path):
    """numbering.xml contains an abstractNum with pStyle back-references."""
    spec = DocumentSpec()
    spec.add_paragraph("Title", heading_level=1)

    with _generate(spec, tmp_path) as z:
        numbering = etree.fromstring(z.read("word/numbering.xml"))
        abstracts = numbering.findall("w:abstractNum", NS)
        assert len(abstracts) == 1  # only heading, no list

        lvls = abstracts[0].findall("w:lvl", NS)
        assert len(lvls) == 4

        expected = [
            ("Heading1", "decimal", "%1"),
            ("Heading2", "decimal", "%1.%2"),
            ("Heading3", "lowerLetter", "(%3)"),
            ("Heading4", "lowerRoman", "(%4)"),
        ]
        for lvl, (p_style, num_fmt, lvl_text) in zip(lvls, expected):
            assert lvl.find("w:pStyle", NS).get(f"{{{NS['w']}}}val") == p_style
            assert lvl.find("w:numFmt", NS).get(f"{{{NS['w']}}}val") == num_fmt
            assert lvl.find("w:lvlText", NS).get(f"{{{NS['w']}}}val") == lvl_text


def test_numbering_num_element(tmp_path):
    """numbering.xml has a <w:num> linking to the heading abstractNum."""
    spec = DocumentSpec()
    spec.add_paragraph("Title", heading_level=1)

    with _generate(spec, tmp_path) as z:
        numbering = etree.fromstring(z.read("word/numbering.xml"))
        nums = numbering.findall("w:num", NS)
        assert len(nums) == 1
        assert nums[0].get(f"{{{NS['w']}}}numId") == "1"
        assert nums[0].find("w:abstractNumId", NS).get(f"{{{NS['w']}}}val") == "0"


# ------------------------------------------------------------------
# Coexistence: list numbering + heading numbering
# ------------------------------------------------------------------


def test_list_and_heading_coexist(tmp_path):
    """Both list and heading numbering can coexist in one document."""
    spec = DocumentSpec()
    spec.add_paragraph("Chapter", heading_level=1)
    spec.add_paragraph("Item 1", numbering=NumberedParagraph(level=0, numbering_id=1))

    with _generate(spec, tmp_path) as z:
        numbering = etree.fromstring(z.read("word/numbering.xml"))
        abstracts = numbering.findall("w:abstractNum", NS)
        assert len(abstracts) == 2  # list + heading

        nums = numbering.findall("w:num", NS)
        assert len(nums) == 2
        num_ids = {n.get(f"{{{NS['w']}}}numId") for n in nums}
        assert num_ids == {"1", "2"}

        # styles.xml heading numId should be "2" since list uses "1"
        styles = etree.fromstring(z.read("word/styles.xml"))
        h1 = styles.find(".//w:style[@w:styleId='Heading1']", NS)
        numId = h1.find("w:pPr/w:numPr/w:numId", NS)
        assert numId.get(f"{{{NS['w']}}}val") == "2"


# ------------------------------------------------------------------
# Combined with comments and tracked changes
# ------------------------------------------------------------------


def test_heading_with_comment(tmp_path):
    """Heading paragraph can also have a comment."""
    spec = DocumentSpec()
    spec.add_paragraph(
        "Important Chapter",
        heading_level=1,
        comments=[Comment(text="Review this", anchor_text="Important")],
    )

    with _generate(spec, tmp_path) as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        para = doc.findall(".//w:p", NS)[0]

        # Has heading style
        pStyle = para.find("w:pPr/w:pStyle", NS)
        assert pStyle.get(f"{{{NS['w']}}}val") == "Heading1"

        # Has comment markers
        assert para.find(".//w:commentRangeStart", NS) is not None


def test_heading_with_tracked_change(tmp_path):
    """Heading paragraph can also have tracked changes."""
    spec = DocumentSpec()
    spec.add_paragraph(
        "Original Title",
        heading_level=2,
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text=" Updated",
                author="Editor",
            )
        ],
    )

    with _generate(spec, tmp_path) as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        para = doc.findall(".//w:p", NS)[0]

        pStyle = para.find("w:pPr/w:pStyle", NS)
        assert pStyle.get(f"{{{NS['w']}}}val") == "Heading2"

        # Has tracked insertion
        assert para.find(".//w:ins", NS) is not None


# ------------------------------------------------------------------
# Validation
# ------------------------------------------------------------------


def test_heading_numbering_validates(tmp_path):
    """Generated heading-numbered document passes validation."""
    spec = DocumentSpec()
    for lvl in range(1, 5):
        spec.add_paragraph(f"Level {lvl} heading", heading_level=lvl)
    spec.add_paragraph("Plain paragraph")

    out = tmp_path / "test.docx"
    DocumentGenerator(spec).generate(out)
    validate_docx(out)  # raises on failure


def test_combined_heading_list_validates(tmp_path):
    """Document with both heading and list numbering passes validation."""
    spec = DocumentSpec()
    spec.add_paragraph("Chapter", heading_level=1)
    spec.add_paragraph("Section", heading_level=2)
    spec.add_paragraph("Item", numbering=NumberedParagraph(level=0, numbering_id=1))
    spec.add_paragraph("Plain text")

    out = tmp_path / "test.docx"
    DocumentGenerator(spec).generate(out)
    validate_docx(out)  # raises on failure
