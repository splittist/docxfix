"""Validation utilities for generated .docx files."""

import zipfile
from pathlib import Path

from lxml import etree


class ValidationError(Exception):
    """Raised when validation fails."""

    pass


class DocumentValidator:
    """Validates .docx files for structural integrity."""

    REQUIRED_FILES = [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
    ]

    def __init__(self, docx_path: str | Path) -> None:
        """Initialize validator with path to .docx file."""
        self.docx_path = Path(docx_path)

    def validate(self) -> None:
        """
        Validate the .docx file.

        Raises:
            ValidationError: If validation fails
        """
        self._validate_zip_structure()
        self._validate_xml_wellformedness()
        self._validate_section_header_footer_integrity()

    def _validate_zip_structure(self) -> None:
        """Validate that required files exist in the ZIP archive."""
        if not self.docx_path.exists():
            raise ValidationError(f"File not found: {self.docx_path}")

        if not zipfile.is_zipfile(self.docx_path):
            raise ValidationError(f"Not a valid ZIP file: {self.docx_path}")

        with zipfile.ZipFile(self.docx_path, "r") as docx_zip:
            available_files = set(docx_zip.namelist())

            for required_file in self.REQUIRED_FILES:
                if required_file not in available_files:
                    raise ValidationError(
                        f"Missing required file: {required_file}"
                    )

    def _validate_xml_wellformedness(self) -> None:
        """Validate that XML files are well-formed."""
        with zipfile.ZipFile(self.docx_path, "r") as docx_zip:
            for filename in docx_zip.namelist():
                if filename.endswith(".xml") or filename.endswith(".rels"):
                    try:
                        content = docx_zip.read(filename)
                        etree.fromstring(content)
                    except etree.XMLSyntaxError as e:
                        raise ValidationError(
                            f"XML syntax error in {filename}: {e}"
                        ) from e

    def _validate_section_header_footer_integrity(self) -> None:
        """Validate section references to header/footer parts and rels."""
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        with zipfile.ZipFile(self.docx_path, "r") as docx_zip:
            document_root = etree.fromstring(docx_zip.read("word/document.xml"))
            rels_root = etree.fromstring(docx_zip.read("word/_rels/document.xml.rels"))
            available_files = set(docx_zip.namelist())

            rel_by_id = {
                rel.get("Id"): rel
                for rel in rels_root.findall(f"{{{rel_ns}}}Relationship")
            }

            section_props = document_root.findall(f".//{{{w_ns}}}sectPr")
            for sect_pr in section_props:
                for tag, expected_type in (
                    ("headerReference", "header"),
                    ("footerReference", "footer"),
                ):
                    for ref in sect_pr.findall(f"{{{w_ns}}}{tag}"):
                        rid = ref.get(f"{{{r_ns}}}id")
                        if not rid:
                            raise ValidationError(f"Section {tag} is missing r:id")
                        if rid not in rel_by_id:
                            raise ValidationError(f"Section {tag} references missing relationship: {rid}")

                        rel = rel_by_id[rid]
                        rel_type = rel.get("Type", "")
                        if not rel_type.endswith(f"/{expected_type}"):
                            raise ValidationError(
                                f"Relationship {rid} type mismatch for {tag}: {rel_type}"
                            )

                        target = rel.get("Target")
                        if not target:
                            raise ValidationError(f"Relationship {rid} has no target")
                        target_path = f"word/{target.lstrip('./')}"
                        if target_path not in available_files:
                            raise ValidationError(
                                f"Missing section part for {rid}: {target_path}"
                            )


def validate_docx(docx_path: str | Path) -> None:
    """
    Validate a .docx file.

    Args:
        docx_path: Path to the .docx file

    Raises:
        ValidationError: If validation fails
    """
    validator = DocumentValidator(docx_path)
    validator.validate()
