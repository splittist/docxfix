"""Document generator for creating .docx files from specifications."""

import io
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from docxfix.spec import ChangeType, DocumentSpec, Paragraph
from docxfix.xml_utils import XMLElement


class DocumentGenerator:
    """Generates .docx files from DocumentSpec."""

    # OOXML namespaces
    NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    }

    def __init__(self, spec: DocumentSpec) -> None:
        """Initialize generator with a document specification."""
        self.spec = spec
        self._revision_counter = 0
        self._comment_counter = 0

    def generate(self, output_path: str | Path) -> None:
        """
        Generate a .docx file at the specified path.

        Args:
            output_path: Path where the .docx file will be created
        """
        output_path = Path(output_path)

        # Create a ZIP file (docx is a ZIP archive)
        with zipfile.ZipFile(
            output_path, "w", zipfile.ZIP_DEFLATED
        ) as docx_zip:
            # Add required files
            docx_zip.writestr("[Content_Types].xml", self._create_content_types())
            docx_zip.writestr("_rels/.rels", self._create_rels())
            docx_zip.writestr(
                "word/_rels/document.xml.rels", self._create_document_rels()
            )
            docx_zip.writestr("word/document.xml", self._create_document())

    def _create_content_types(self) -> bytes:
        """Create [Content_Types].xml."""
        types = etree.Element(
            "Types",
            xmlns="http://schemas.openxmlformats.org/package/2006/content-types",
        )
        etree.SubElement(
            types,
            "Default",
            Extension="rels",
            ContentType="application/vnd.openxmlformats-package.relationships+xml",
        )
        etree.SubElement(
            types,
            "Default",
            Extension="xml",
            ContentType="application/xml",
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/document.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document.main+xml"
            ),
        )
        return etree.tostring(
            types, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_rels(self) -> bytes:
        """Create _rels/.rels."""
        rels = etree.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId1",
            Type=(
                "http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships/officeDocument"
            ),
            Target="word/document.xml",
        )
        return etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_document_rels(self) -> bytes:
        """Create word/_rels/document.xml.rels."""
        rels = etree.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )
        return etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_document(self) -> bytes:
        """Create word/document.xml with paragraphs and features."""
        document = etree.Element(
            f"{{{self.NAMESPACES['w']}}}document",
            nsmap=self.NAMESPACES,
        )
        body = etree.SubElement(document, f"{{{self.NAMESPACES['w']}}}body")

        # Add each paragraph
        for para_spec in self.spec.paragraphs:
            self._add_paragraph(body, para_spec)

        return etree.tostring(
            document, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _add_paragraph(self, body: XMLElement, para_spec: Paragraph) -> None:
        """Add a paragraph to the body."""
        w_ns = self.NAMESPACES["w"]
        para = etree.SubElement(body, f"{{{w_ns}}}p")

        # Add tracked changes if present
        if para_spec.tracked_changes:
            for change in para_spec.tracked_changes:
                self._add_tracked_change(para, change)
        else:
            # Simple run with text
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = para_spec.text

    def _add_tracked_change(
        self, para: XMLElement, change
    ) -> None:  # change: TrackedChange
        """Add a tracked change to a paragraph."""
        w_ns = self.NAMESPACES["w"]
        self._revision_counter += 1

        # Format date
        date_str = change.date.strftime("%Y-%m-%dT%H:%M:%SZ")

        if change.change_type == ChangeType.INSERTION:
            # Create insertion element
            ins = etree.SubElement(
                para,
                f"{{{w_ns}}}ins",
                {
                    f"{{{w_ns}}}id": str(self._revision_counter),
                    f"{{{w_ns}}}author": change.author,
                    f"{{{w_ns}}}date": date_str,
                },
            )
            # Add run with text inside insertion
            run = etree.SubElement(ins, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = change.text

        elif change.change_type == ChangeType.DELETION:
            # Create deletion element
            delete = etree.SubElement(
                para,
                f"{{{w_ns}}}del",
                {
                    f"{{{w_ns}}}id": str(self._revision_counter),
                    f"{{{w_ns}}}author": change.author,
                    f"{{{w_ns}}}date": date_str,
                },
            )
            # Add run with deleted text inside deletion
            run = etree.SubElement(delete, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}delText")
            text_elem.text = change.text
