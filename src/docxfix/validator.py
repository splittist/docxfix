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
        """Validate the .docx file.

        Raises:
            ValidationError: If validation fails
        """
        self._validate_zip_structure()
        self._validate_xml_wellformedness()
        self._validate_section_header_footer_integrity()
        self._validate_comment_id_uniqueness()
        self._validate_tracked_change_id_uniqueness()
        self._validate_comment_anchor_integrity()
        self._validate_relationship_completeness()
        self._validate_content_type_coverage()

    def _validate_zip_structure(self) -> None:
        """Validate that required files exist in the ZIP."""
        if not self.docx_path.exists():
            raise ValidationError(
                f"File not found: {self.docx_path}"
            )

        if not zipfile.is_zipfile(self.docx_path):
            raise ValidationError(
                f"Not a valid ZIP file: {self.docx_path}"
            )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            available = set(docx_zip.namelist())
            for req in self.REQUIRED_FILES:
                if req not in available:
                    raise ValidationError(
                        f"Missing required file: {req}"
                    )

    def _validate_xml_wellformedness(self) -> None:
        """Validate that XML files are well-formed."""
        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            for filename in docx_zip.namelist():
                if filename.endswith(
                    ".xml"
                ) or filename.endswith(".rels"):
                    try:
                        content = docx_zip.read(filename)
                        etree.fromstring(content)
                    except etree.XMLSyntaxError as e:
                        raise ValidationError(
                            f"XML syntax error in"
                            f" {filename}: {e}"
                        ) from e

    def _validate_section_header_footer_integrity(
        self,
    ) -> None:
        """Validate section references to header/footer parts."""
        w_ns = (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )
        r_ns = (
            "http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships"
        )
        rel_ns = (
            "http://schemas.openxmlformats.org/"
            "package/2006/relationships"
        )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            doc_root = etree.fromstring(
                docx_zip.read("word/document.xml")
            )
            rels_root = etree.fromstring(
                docx_zip.read(
                    "word/_rels/document.xml.rels"
                )
            )
            available = set(docx_zip.namelist())

            rel_by_id = {
                rel.get("Id"): rel
                for rel in rels_root.findall(
                    f"{{{rel_ns}}}Relationship"
                )
            }

            for sect_pr in doc_root.findall(
                f".//{{{w_ns}}}sectPr"
            ):
                for tag, expected in (
                    ("headerReference", "header"),
                    ("footerReference", "footer"),
                ):
                    for ref in sect_pr.findall(
                        f"{{{w_ns}}}{tag}"
                    ):
                        rid = ref.get(f"{{{r_ns}}}id")
                        if not rid:
                            raise ValidationError(
                                f"Section {tag}"
                                " is missing r:id"
                            )
                        if rid not in rel_by_id:
                            raise ValidationError(
                                f"Section {tag} references"
                                f" missing relationship:"
                                f" {rid}"
                            )

                        rel = rel_by_id[rid]
                        rel_type = rel.get("Type", "")
                        if not rel_type.endswith(
                            f"/{expected}"
                        ):
                            raise ValidationError(
                                f"Relationship {rid} type"
                                f" mismatch for {tag}:"
                                f" {rel_type}"
                            )

                        target = rel.get("Target")
                        if not target:
                            raise ValidationError(
                                f"Relationship {rid}"
                                " has no target"
                            )
                        target_path = (
                            f"word/{target.lstrip('./')}"
                        )
                        if target_path not in available:
                            raise ValidationError(
                                f"Missing section part for"
                                f" {rid}: {target_path}"
                            )

    def _validate_comment_id_uniqueness(self) -> None:
        """Check that comment IDs in comments.xml are unique."""
        w_ns = (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            if "word/comments.xml" not in docx_zip.namelist():
                return

            root = etree.fromstring(
                docx_zip.read("word/comments.xml")
            )
            ids: list[str] = []
            for comment in root.findall(
                f"{{{w_ns}}}comment"
            ):
                cid = comment.get(f"{{{w_ns}}}id")
                if cid is not None:
                    ids.append(cid)

            dupes = {x for x in ids if ids.count(x) > 1}
            if dupes:
                raise ValidationError(
                    f"Duplicate comment IDs: {dupes}"
                )

    def _validate_tracked_change_id_uniqueness(
        self,
    ) -> None:
        """Check that tracked change IDs are unique."""
        w_ns = (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            root = etree.fromstring(
                docx_zip.read("word/document.xml")
            )

            ids: list[str] = []
            for tag in ("ins", "del"):
                for elem in root.findall(
                    f".//{{{w_ns}}}{tag}"
                ):
                    tid = elem.get(f"{{{w_ns}}}id")
                    if tid is not None:
                        ids.append(tid)

            dupes = {x for x in ids if ids.count(x) > 1}
            if dupes:
                raise ValidationError(
                    "Duplicate tracked change IDs:"
                    f" {dupes}"
                )

    def _validate_comment_anchor_integrity(self) -> None:
        """Verify commentRangeStart/End pairs match."""
        w_ns = (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            root = etree.fromstring(
                docx_zip.read("word/document.xml")
            )

            starts = {
                elem.get(f"{{{w_ns}}}id")
                for elem in root.findall(
                    f".//{{{w_ns}}}commentRangeStart"
                )
                if elem.get(f"{{{w_ns}}}id") is not None
            }
            ends = {
                elem.get(f"{{{w_ns}}}id")
                for elem in root.findall(
                    f".//{{{w_ns}}}commentRangeEnd"
                )
                if elem.get(f"{{{w_ns}}}id") is not None
            }

            if starts and starts != ends:
                unmatched_starts = starts - ends
                unmatched_ends = ends - starts
                parts = []
                if unmatched_starts:
                    parts.append(
                        "commentRangeStart without End:"
                        f" {unmatched_starts}"
                    )
                if unmatched_ends:
                    parts.append(
                        "commentRangeEnd without Start:"
                        f" {unmatched_ends}"
                    )
                raise ValidationError(
                    "Comment anchor mismatch: "
                    + "; ".join(parts)
                )

    def _validate_relationship_completeness(
        self,
    ) -> None:
        """Verify every rId in document.xml exists in rels."""
        r_ns = (
            "http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships"
        )
        rel_ns = (
            "http://schemas.openxmlformats.org/"
            "package/2006/relationships"
        )

        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            if (
                "word/_rels/document.xml.rels"
                not in docx_zip.namelist()
            ):
                return

            doc_root = etree.fromstring(
                docx_zip.read("word/document.xml")
            )
            rels_root = etree.fromstring(
                docx_zip.read(
                    "word/_rels/document.xml.rels"
                )
            )

            defined_ids = {
                rel.get("Id")
                for rel in rels_root.findall(
                    f"{{{rel_ns}}}Relationship"
                )
                if rel.get("Id") is not None
            }

            # Find all r:id references in document.xml
            referenced_ids: set[str] = set()
            for elem in doc_root.iter():
                rid = elem.get(f"{{{r_ns}}}id")
                if rid is not None:
                    referenced_ids.add(rid)

            missing = referenced_ids - defined_ids
            if missing:
                raise ValidationError(
                    "document.xml references undefined"
                    f" relationships: {missing}"
                )

    def _validate_content_type_coverage(self) -> None:
        """Verify every ZIP part has a content type."""
        with zipfile.ZipFile(
            self.docx_path, "r"
        ) as docx_zip:
            ct_root = etree.fromstring(
                docx_zip.read("[Content_Types].xml")
            )

            ct_ns = (
                "http://schemas.openxmlformats.org/"
                "package/2006/content-types"
            )

            # Collect Default extensions
            default_exts = {
                d.get("Extension")
                for d in ct_root.findall(
                    f"{{{ct_ns}}}Default"
                )
                if d.get("Extension") is not None
            }

            # Collect Override part names
            override_parts = {
                o.get("PartName", "").lstrip("/")
                for o in ct_root.findall(
                    f"{{{ct_ns}}}Override"
                )
            }

            for part_name in docx_zip.namelist():
                # Skip rels directory files — they use
                # the .rels Default extension
                ext = part_name.rsplit(".", 1)[-1]
                if ext in default_exts:
                    continue

                if part_name not in override_parts:
                    raise ValidationError(
                        f"Part {part_name} has no"
                        " matching content type"
                    )


def validate_docx(docx_path: str | Path) -> None:
    """Validate a .docx file.

    Args:
        docx_path: Path to the .docx file

    Raises:
        ValidationError: If validation fails
    """
    validator = DocumentValidator(docx_path)
    validator.validate()
