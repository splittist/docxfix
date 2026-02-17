"""Document generator for creating .docx files from specifications."""

import datetime as dt
import random
import zipfile
from pathlib import Path

from lxml import etree

from docxfix.boilerplate import (
    create_app_properties,
    create_core_properties,
    create_endnotes,
    create_font_table,
    create_footnotes,
    create_settings,
    create_theme,
    create_web_settings,
)
from docxfix.constants import NAMESPACES, WORD_NAMESPACES
from docxfix.parts.comments import (
    add_paragraph_with_comments,
    create_comments,
    create_comments_extended,
)
from docxfix.parts.context import GeneratorContext
from docxfix.parts.numbering import create_numbering
from docxfix.parts.sections import (
    add_section_properties,
    build_section_part_manifest,
    create_header_footer_part,
    normalize_sections,
)
from docxfix.parts.styles import create_styles
from docxfix.parts.tracked_changes import (
    add_paragraph_with_comments_and_tracked_changes,
    add_paragraph_with_tracked_changes,
)
from docxfix.spec import (
    DocumentSpec,
    Paragraph,
    SectionSpec,
)
from docxfix.xml_utils import XMLElement


class DocumentGenerator:
    """Generates .docx files from DocumentSpec."""

    # Fixed reference datetime used when seed is set.
    _REFERENCE_DATETIME = dt.datetime(
        2024, 1, 1, tzinfo=dt.UTC
    )

    def __init__(self, spec: DocumentSpec) -> None:
        """Initialize generator with a document specification."""
        self.spec = spec

        # Build a seeded or unseeded RNG instance
        rng = random.Random(spec.seed)
        self._ctx = GeneratorContext(
            namespaces=NAMESPACES,
            word_namespaces=WORD_NAMESPACES,
            _rng=rng,
        )
        self._section_layout = normalize_sections(
            self.spec.sections, len(self.spec.paragraphs)
        )
        self._section_manifest, self._section_refs = (
            build_section_part_manifest(self._section_layout)
        )

        # When seeded, fix datetimes on spec objects that
        # used defaults so output is fully deterministic.
        if spec.seed is not None:
            self._fix_spec_datetimes()

    def _fix_spec_datetimes(self) -> None:
        """Replace non-deterministic datetimes with a fixed reference."""
        ref = self._REFERENCE_DATETIME
        for para in self.spec.paragraphs:
            for tc in para.tracked_changes:
                tc.date = ref
            for comment in para.comments:
                comment.date = ref
                for reply in comment.replies:
                    reply.date = ref

    def generate(self, output_path: str | Path) -> None:
        """Generate a .docx file at the specified path."""
        output_path = Path(output_path)

        has_comments = any(
            p.comments for p in self.spec.paragraphs
        )
        has_numbering = any(
            p.numbering for p in self.spec.paragraphs
        )
        has_heading_numbering = any(
            p.heading_level for p in self.spec.paragraphs
        )
        needs_numbering = has_numbering or has_heading_numbering

        # Build rels (populates _section_refs with rIds)
        doc_rels = self._create_document_rels(
            has_comments, needs_numbering
        )

        with zipfile.ZipFile(
            output_path, "w", zipfile.ZIP_DEFLATED
        ) as docx_zip:
            docx_zip.writestr(
                "[Content_Types].xml",
                self._create_content_types(
                    has_comments, needs_numbering
                ),
            )
            docx_zip.writestr(
                "_rels/.rels", self._create_rels()
            )
            docx_zip.writestr(
                "word/_rels/document.xml.rels", doc_rels
            )
            docx_zip.writestr(
                "word/document.xml",
                self._create_document(),
            )
            docx_zip.writestr(
                "word/settings.xml",
                create_settings(self._section_layout),
            )
            docx_zip.writestr(
                "word/webSettings.xml",
                create_web_settings(),
            )
            docx_zip.writestr(
                "word/footnotes.xml",
                create_footnotes(
                    self._ctx.generate_hex_id
                ),
            )
            docx_zip.writestr(
                "word/endnotes.xml",
                create_endnotes(
                    self._ctx.generate_hex_id
                ),
            )
            docx_zip.writestr(
                "word/fontTable.xml", create_font_table()
            )
            docx_zip.writestr(
                "word/theme/theme1.xml", create_theme()
            )
            docx_zip.writestr(
                "docProps/core.xml",
                create_core_properties(
                    self.spec.title,
                    self.spec.author,
                    timestamp=(
                        self._REFERENCE_DATETIME
                        if self.spec.seed is not None
                        else None
                    ),
                ),
            )
            docx_zip.writestr(
                "docProps/app.xml",
                create_app_properties(),
            )

            for part in self._section_manifest:
                docx_zip.writestr(
                    part["path"],
                    create_header_footer_part(
                        part["kind"], part["text"]
                    ),
                )

            if has_comments:
                docx_zip.writestr(
                    "word/comments.xml",
                    create_comments(self._ctx),
                )
                docx_zip.writestr(
                    "word/commentsExtended.xml",
                    create_comments_extended(self._ctx),
                )

            if needs_numbering:
                docx_zip.writestr(
                    "word/numbering.xml",
                    create_numbering(
                        has_numbering,
                        has_heading_numbering,
                        NAMESPACES,
                        WORD_NAMESPACES,
                    ),
                )
                docx_zip.writestr(
                    "word/styles.xml",
                    create_styles(
                        has_heading_numbering,
                        self.spec.paragraphs,
                        NAMESPACES,
                        WORD_NAMESPACES,
                    ),
                )

    def _create_content_types(
        self,
        has_comments: bool = False,
        has_numbering: bool = False,
    ) -> bytes:
        """Create [Content_Types].xml."""
        types = etree.Element(
            "Types",
            xmlns=(
                "http://schemas.openxmlformats.org/"
                "package/2006/content-types"
            ),
        )
        etree.SubElement(
            types,
            "Default",
            Extension="rels",
            ContentType=(
                "application/vnd.openxmlformats-package"
                ".relationships+xml"
            ),
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

        if has_comments:
            for part_name, suffix in (
                ("/word/comments.xml", "comments"),
                (
                    "/word/commentsExtended.xml",
                    "commentsExtended",
                ),
            ):
                etree.SubElement(
                    types,
                    "Override",
                    PartName=part_name,
                    ContentType=(
                        "application/vnd.openxmlformats"
                        "-officedocument.wordprocessingml."
                        f"{suffix}+xml"
                    ),
                )

        if has_numbering:
            for part_name, suffix in (
                ("/word/numbering.xml", "numbering"),
                ("/word/styles.xml", "styles"),
            ):
                etree.SubElement(
                    types,
                    "Override",
                    PartName=part_name,
                    ContentType=(
                        "application/vnd.openxmlformats"
                        "-officedocument.wordprocessingml."
                        f"{suffix}+xml"
                    ),
                )

        for part in self._section_manifest:
            etree.SubElement(
                types,
                "Override",
                PartName=(
                    f"/word/{part['path'].split('/', 1)[1]}"
                ),
                ContentType=part["content_type"],
            )

        # Standard parts
        standard_parts = [
            ("/word/settings.xml", "settings"),
            ("/word/webSettings.xml", "webSettings"),
            ("/word/footnotes.xml", "footnotes"),
            ("/word/endnotes.xml", "endnotes"),
            ("/word/fontTable.xml", "fontTable"),
        ]
        for part_name, suffix in standard_parts:
            etree.SubElement(
                types,
                "Override",
                PartName=part_name,
                ContentType=(
                    "application/vnd.openxmlformats"
                    "-officedocument.wordprocessingml."
                    f"{suffix}+xml"
                ),
            )

        etree.SubElement(
            types,
            "Override",
            PartName="/word/theme/theme1.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "theme+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/docProps/core.xml",
            ContentType=(
                "application/vnd.openxmlformats-package"
                ".core-properties+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/docProps/app.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "extended-properties+xml"
            ),
        )

        return etree.tostring(
            types,
            xml_declaration=True,
            encoding="UTF-8",
            pretty_print=True,
        )

    def _create_rels(self) -> bytes:
        """Create _rels/.rels."""
        rels = etree.Element(
            "Relationships",
            xmlns=(
                "http://schemas.openxmlformats.org/"
                "package/2006/relationships"
            ),
        )
        rel_base = (
            "http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships"
        )
        pkg_base = (
            "http://schemas.openxmlformats.org/"
            "package/2006/relationships"
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId1",
            Type=f"{rel_base}/officeDocument",
            Target="word/document.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId2",
            Type=(
                f"{pkg_base}/metadata/core-properties"
            ),
            Target="docProps/core.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId3",
            Type=f"{rel_base}/extended-properties",
            Target="docProps/app.xml",
        )
        return etree.tostring(
            rels,
            xml_declaration=True,
            encoding="UTF-8",
            pretty_print=True,
        )

    def _create_document_rels(
        self,
        has_comments: bool = False,
        has_numbering: bool = False,
    ) -> bytes:
        """Create word/_rels/document.xml.rels."""
        rel_base = (
            "http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships"
        )
        rels = etree.Element(
            "Relationships",
            xmlns=(
                "http://schemas.openxmlformats.org/"
                "package/2006/relationships"
            ),
        )

        next_id = 1

        if has_comments:
            comment_rels = [
                (f"{rel_base}/comments", "comments.xml"),
                (
                    "http://schemas.microsoft.com/office/"
                    "2011/relationships/commentsExtended",
                    "commentsExtended.xml",
                ),
            ]
            for rel_type, target in comment_rels:
                etree.SubElement(
                    rels,
                    "Relationship",
                    Id=f"rId{next_id}",
                    Type=rel_type,
                    Target=target,
                )
                next_id += 1

        if has_numbering:
            for suffix in ("numbering", "styles"):
                etree.SubElement(
                    rels,
                    "Relationship",
                    Id=f"rId{next_id}",
                    Type=f"{rel_base}/{suffix}",
                    Target=f"{suffix}.xml",
                )
                next_id += 1

        for part in self._section_manifest:
            rid = f"rId{next_id}"
            next_id += 1
            etree.SubElement(
                rels,
                "Relationship",
                Id=rid,
                Type=part["relationship_type"],
                Target=part["target"],
            )
            si = part["section_index"]
            self._section_refs[si][part["kind"]][
                part["variant"]
            ] = rid

        for suffix in (
            "settings",
            "webSettings",
            "footnotes",
            "endnotes",
            "fontTable",
        ):
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type=f"{rel_base}/{suffix}",
                Target=f"{suffix}.xml",
            )
            next_id += 1

        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id}",
            Type=f"{rel_base}/theme",
            Target="theme/theme1.xml",
        )

        return etree.tostring(
            rels,
            xml_declaration=True,
            encoding="UTF-8",
            pretty_print=True,
        )

    def _create_document(self) -> bytes:
        """Create word/document.xml with paragraphs."""
        document = etree.Element(
            f"{{{NAMESPACES['w']}}}document",
            nsmap=WORD_NAMESPACES,
        )
        document.set(
            f"{{{NAMESPACES['mc']}}}Ignorable",
            "w14 w15 w16se w16cid w16 w16cex"
            " w16sdtdh w16sdtfl w16du wp14",
        )
        body = etree.SubElement(
            document, f"{{{NAMESPACES['w']}}}body"
        )

        sections = self._section_layout
        paragraph_count = len(self.spec.paragraphs)
        section_starts = [
            s.start_paragraph for s in sections
        ] + [paragraph_count]

        boundary_to_section: dict[int, SectionSpec] = {}
        for idx in range(len(sections) - 1):
            boundary_to_section[
                section_starts[idx + 1] - 1
            ] = sections[idx]

        for para_index, para_spec in enumerate(
            self.spec.paragraphs
        ):
            para = self._add_paragraph(body, para_spec)
            if para_index in boundary_to_section:
                add_section_properties(
                    para,
                    boundary_to_section[para_index],
                    self._section_layout,
                    self._section_refs,
                    is_body_level=False,
                )

        add_section_properties(
            body,
            sections[-1],
            self._section_layout,
            self._section_refs,
            is_body_level=True,
        )

        return etree.tostring(
            document,
            xml_declaration=True,
            encoding="UTF-8",
            pretty_print=True,
        )

    def _add_paragraph(
        self, body: XMLElement, para_spec: Paragraph
    ) -> XMLElement:
        """Add a paragraph to the body."""
        w_ns = NAMESPACES["w"]
        w14_ns = NAMESPACES["w14"]
        para = etree.SubElement(body, f"{{{w_ns}}}p")

        para_id = self._ctx.generate_hex_id(8)
        para.set(f"{{{w14_ns}}}paraId", para_id)
        para.set(f"{{{w14_ns}}}textId", "77777777")

        if para_spec.numbering:
            p_pr = etree.SubElement(
                para, f"{{{w_ns}}}pPr"
            )
            p_style = etree.SubElement(
                p_pr, f"{{{w_ns}}}pStyle"
            )
            p_style.set(
                f"{{{w_ns}}}val", "ListParagraph"
            )
            num_pr = etree.SubElement(
                p_pr, f"{{{w_ns}}}numPr"
            )
            ilvl = etree.SubElement(
                num_pr, f"{{{w_ns}}}ilvl"
            )
            ilvl.set(
                f"{{{w_ns}}}val",
                str(para_spec.numbering.level),
            )
            num_id = etree.SubElement(
                num_pr, f"{{{w_ns}}}numId"
            )
            num_id.set(
                f"{{{w_ns}}}val",
                str(para_spec.numbering.numbering_id),
            )
        elif para_spec.heading_level:
            p_pr = etree.SubElement(
                para, f"{{{w_ns}}}pPr"
            )
            p_style = etree.SubElement(
                p_pr, f"{{{w_ns}}}pStyle"
            )
            p_style.set(
                f"{{{w_ns}}}val",
                f"Heading{para_spec.heading_level}",
            )

        if (
            para_spec.comments
            and para_spec.tracked_changes
        ):
            add_paragraph_with_comments_and_tracked_changes(
                para, para_spec, self._ctx
            )
        elif para_spec.comments:
            add_paragraph_with_comments(
                para, para_spec, self._ctx
            )
        elif para_spec.tracked_changes:
            add_paragraph_with_tracked_changes(
                para, para_spec, self._ctx
            )
        else:
            run = etree.SubElement(
                para, f"{{{w_ns}}}r"
            )
            text_elem = etree.SubElement(
                run, f"{{{w_ns}}}t"
            )
            text_elem.text = para_spec.text

        return para

    # Keep _generate_hex_id as a method for backward compat
    def _generate_hex_id(self, length: int = 8) -> str:
        """Generate a random hexadecimal ID."""
        return self._ctx.generate_hex_id(length)
