"""Integration tests for section generation."""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec, HeaderFooterSet, PageOrientation
from docxfix.validator import DocumentValidator, ValidationError


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def test_multi_section_portrait_to_landscape_emits_sectpr():
    """Section boundaries should create paragraph-level and body-level sectPr elements."""
    spec = DocumentSpec()
    spec.add_paragraph("Section 1 para")
    spec.add_paragraph("Section 2 para")
    spec.add_section(1, orientation=PageOrientation.LANDSCAPE)

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "sections.docx"
        DocumentGenerator(spec).generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            root = etree.fromstring(docx_zip.read("word/document.xml"))
            sectprs = root.findall(f".//{{{W_NS}}}sectPr")
            assert len(sectprs) == 2

            first_pg_sz = sectprs[0].find(f"{{{W_NS}}}pgSz")
            assert first_pg_sz is not None
            assert first_pg_sz.get(f"{{{W_NS}}}orient") is None

            second_pg_sz = sectprs[1].find(f"{{{W_NS}}}pgSz")
            assert second_pg_sz is not None
            assert second_pg_sz.get(f"{{{W_NS}}}orient") == "landscape"


def test_section_specific_header_footer_variants_packaged_and_wired():
    """Section header/footer variants should be emitted with rel wiring and content-types."""
    spec = DocumentSpec()
    spec.add_paragraph("First section text")
    spec.add_paragraph("Second section text")
    spec.sections[0].headers = HeaderFooterSet(default="S1 default header")
    spec.add_section(
        1,
        headers=HeaderFooterSet(default="S2 default", first="S2 first", even="S2 even"),
        footers=HeaderFooterSet(default="S2 footer", even="S2 even footer"),
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "header-footer-sections.docx"
        DocumentGenerator(spec).generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            files = set(docx_zip.namelist())
            assert any(name.startswith("word/header") for name in files)
            assert any(name.startswith("word/footer") for name in files)

            content_types = etree.fromstring(docx_zip.read("[Content_Types].xml"))
            overrides = content_types.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
            override_parts = {override.get("PartName") for override in overrides}
            assert any(part and part.startswith("/word/header") for part in override_parts)
            assert any(part and part.startswith("/word/footer") for part in override_parts)

            document = etree.fromstring(docx_zip.read("word/document.xml"))
            sectprs = document.findall(f".//{{{W_NS}}}sectPr")
            assert len(sectprs) == 2

            final = sectprs[-1]
            header_types = {
                ref.get(f"{{{W_NS}}}type")
                for ref in final.findall(f"{{{W_NS}}}headerReference")
            }
            footer_types = {
                ref.get(f"{{{W_NS}}}type")
                for ref in final.findall(f"{{{W_NS}}}footerReference")
            }
            assert {"default", "first", "even"}.issubset(header_types)
            assert {"default", "even"}.issubset(footer_types)
            assert final.find(f"{{{W_NS}}}titlePg") is not None

            settings = etree.fromstring(docx_zip.read("word/settings.xml"))
            assert settings.find(f"{{{W_NS}}}evenAndOddHeaders") is not None


def test_validator_fails_when_section_relationship_missing_target():
    """Semantic validator should reject broken section header/footer targets."""
    spec = DocumentSpec()
    spec.add_paragraph("Only paragraph")
    spec.sections[0].headers = HeaderFooterSet(default="Header")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "invalid-section.docx"
        DocumentGenerator(spec).generate(output_path)

        # Corrupt relationship target
        with zipfile.ZipFile(output_path, "a") as docx_zip:
            rels_root = etree.fromstring(docx_zip.read("word/_rels/document.xml.rels"))
            rel = rels_root.find("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            assert rel is not None
            rel.set("Target", "missing-header.xml")
            docx_zip.writestr(
                "word/_rels/document.xml.rels",
                etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8"),
            )

        validator = DocumentValidator(output_path)
        try:
            validator.validate()
        except ValidationError as err:
            assert "Missing section part" in str(err)
        else:
            raise AssertionError("Expected ValidationError for missing section target")


def test_ppr_is_first_child_of_boundary_paragraph():
    """pPr with sectPr must be first child of boundary paragraph."""
    spec = DocumentSpec()
    spec.add_paragraph("Section 1 text")
    spec.add_paragraph("Section 2 text")
    spec.add_section(1, orientation=PageOrientation.LANDSCAPE)

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "ppr-order.docx"
        DocumentGenerator(spec).generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            root = etree.fromstring(docx_zip.read("word/document.xml"))
            body = root.find(f"{{{W_NS}}}body")
            paragraphs = body.findall(f"{{{W_NS}}}p")

            # Find the boundary paragraph (has pPr/sectPr)
            boundary = None
            for p in paragraphs:
                p_pr = p.find(f"{{{W_NS}}}pPr")
                if p_pr is not None and p_pr.find(f"{{{W_NS}}}sectPr") is not None:
                    boundary = p
                    break

            assert boundary is not None, "No boundary paragraph with pPr/sectPr found"
            first_child = boundary[0]
            assert first_child.tag == f"{{{W_NS}}}pPr", (
                f"pPr must be first child of boundary paragraph, got {first_child.tag}"
            )


def test_header_footer_parts_have_pstyle_and_mc_ignorable():
    """Header/footer parts should have pStyle and mc:Ignorable attribute."""
    spec = DocumentSpec()
    spec.add_paragraph("Text")
    spec.sections[0].headers = HeaderFooterSet(default="My Header")
    spec.sections[0].footers = HeaderFooterSet(default="My Footer")

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "hf-style.docx"
        DocumentGenerator(spec).generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            names = docx_zip.namelist()
            header_files = [n for n in names if n.startswith("word/header")]
            footer_files = [n for n in names if n.startswith("word/footer")]

            assert len(header_files) >= 1
            assert len(footer_files) >= 1

            # Check header
            hdr_root = etree.fromstring(docx_zip.read(header_files[0]))
            assert hdr_root.get(f"{{{MC_NS}}}Ignorable") is not None
            hdr_pstyle = hdr_root.find(f".//{{{W_NS}}}pStyle")
            assert hdr_pstyle is not None
            assert hdr_pstyle.get(f"{{{W_NS}}}val") == "Header"

            # Check footer
            ftr_root = etree.fromstring(docx_zip.read(footer_files[0]))
            assert ftr_root.get(f"{{{MC_NS}}}Ignorable") is not None
            ftr_pstyle = ftr_root.find(f".//{{{W_NS}}}pStyle")
            assert ftr_pstyle is not None
            assert ftr_pstyle.get(f"{{{W_NS}}}val") == "Footer"


def test_three_section_corpus_matching():
    """Generate a 3-section doc matching corpus/sections.md and validate structure."""
    spec = DocumentSpec()

    # Section 1: Title page
    spec.add_paragraph("Title Page")

    # Section 2: Body with lorem text (multiple paragraphs)
    for i in range(10):
        spec.add_paragraph(f"Lorem ipsum paragraph {i + 1} of the body section.")

    # Section 3: Landscape page
    spec.add_paragraph("Section 3, which is landscape")

    # Configure sections: section 2 starts at paragraph 1, section 3 at paragraph 11
    spec.add_section(
        1,
        headers=HeaderFooterSet(default="Header for Section 2"),
        footers=HeaderFooterSet(default="Footer Section 2"),
        page_number_start=1,
    )
    spec.add_section(
        11,
        orientation=PageOrientation.LANDSCAPE,
        headers=HeaderFooterSet(default="Header for section 3"),
        footers=HeaderFooterSet(default="Footer Section 3"),
        page_number_start=1,
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "three-sections.docx"
        DocumentGenerator(spec).generate(output_path)

        with zipfile.ZipFile(output_path, "r") as docx_zip:
            root = etree.fromstring(docx_zip.read("word/document.xml"))
            sectprs = root.findall(f".//{{{W_NS}}}sectPr")

            # Three sectPr elements total
            assert len(sectprs) == 3, f"Expected 3 sectPr elements, got {len(sectprs)}"

            # First two are within pPr (paragraph-level)
            for i in range(2):
                parent = sectprs[i].getparent()
                assert parent.tag == f"{{{W_NS}}}pPr", (
                    f"sectPr[{i}] parent should be pPr, got {parent.tag}"
                )
                grandparent = parent.getparent()
                assert grandparent.tag == f"{{{W_NS}}}p", (
                    f"sectPr[{i}] grandparent should be p, got {grandparent.tag}"
                )
                # pPr must be first child
                assert grandparent[0].tag == f"{{{W_NS}}}pPr"

            # Last sectPr is body-level (direct child of body)
            assert sectprs[2].getparent().tag == f"{{{W_NS}}}body"

            # First two sections: portrait
            for i in range(2):
                pg_sz = sectprs[i].find(f"{{{W_NS}}}pgSz")
                assert pg_sz is not None
                w_val = int(pg_sz.get(f"{{{W_NS}}}w"))
                h_val = int(pg_sz.get(f"{{{W_NS}}}h"))
                assert w_val < h_val, f"sectPr[{i}] should be portrait (w < h)"

            # Third section: landscape
            pg_sz = sectprs[2].find(f"{{{W_NS}}}pgSz")
            w_val = int(pg_sz.get(f"{{{W_NS}}}w"))
            h_val = int(pg_sz.get(f"{{{W_NS}}}h"))
            assert w_val > h_val, "Last sectPr should be landscape (w > h)"

            # Sections 2 and 3 have header/footer references
            for i in [1, 2]:
                hdr_refs = sectprs[i].findall(f"{{{W_NS}}}headerReference")
                ftr_refs = sectprs[i].findall(f"{{{W_NS}}}footerReference")
                assert len(hdr_refs) >= 1, f"sectPr[{i}] needs headerReference"
                assert len(ftr_refs) >= 1, f"sectPr[{i}] needs footerReference"
                assert any(
                    r.get(f"{{{W_NS}}}type") == "default" for r in hdr_refs
                ), f"sectPr[{i}] needs default headerReference"
                assert any(
                    r.get(f"{{{W_NS}}}type") == "default" for r in ftr_refs
                ), f"sectPr[{i}] needs default footerReference"

            # Sections 2 and 3 have pgNumType with start="1"
            for i in [1, 2]:
                pg_num = sectprs[i].find(f"{{{W_NS}}}pgNumType")
                assert pg_num is not None, f"sectPr[{i}] needs pgNumType"
                assert pg_num.get(f"{{{W_NS}}}start") == "1"

            # Verify header/footer parts exist in the ZIP
            all_names = docx_zip.namelist()
            header_parts = [n for n in all_names if n.startswith("word/header")]
            footer_parts = [n for n in all_names if n.startswith("word/footer")]
            assert len(header_parts) >= 2, "Need at least 2 header parts"
            assert len(footer_parts) >= 2, "Need at least 2 footer parts"
