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
