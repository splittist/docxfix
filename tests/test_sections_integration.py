"""Integration tests for section generation."""

import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec, HeaderFooterSet, PageOrientation
from docxfix.validator import DocumentValidator, ValidationError


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
