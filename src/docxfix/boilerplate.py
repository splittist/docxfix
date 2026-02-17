"""Stateless boilerplate XML part generators for .docx files."""

import base64
import datetime as dt

from lxml import etree

from docxfix.constants import (
    APP_PROPERTIES_XML,
    FONT_TABLE_XML,
    NAMESPACES,
    SETTINGS_XML,
    THEME_XML_B64,
    WEB_SETTINGS_XML,
)
from docxfix.spec import SectionSpec
from docxfix.xml_utils import XMLElement


def create_settings(section_layout: list[SectionSpec]) -> bytes:
    """Create word/settings.xml."""
    root = etree.fromstring(SETTINGS_XML.encode("utf-8"))
    w_ns = NAMESPACES["w"]
    has_even = any(
        section.headers.even is not None or section.footers.even is not None
        for section in section_layout
    )
    if has_even and root.find(f"{{{w_ns}}}evenAndOddHeaders") is None:
        root.insert(0, etree.Element(f"{{{w_ns}}}evenAndOddHeaders"))
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", pretty_print=False
    )


def create_web_settings() -> bytes:
    """Create word/webSettings.xml."""
    return WEB_SETTINGS_XML.encode("utf-8")


def create_footnotes(generate_hex_id) -> bytes:
    """Create word/footnotes.xml."""
    w_ns = NAMESPACES["w"]
    w14_ns = NAMESPACES["w14"]
    footnotes = etree.Element(
        f"{{{w_ns}}}footnotes",
        nsmap={"w": w_ns, "w14": w14_ns},
    )
    add_note_separator(
        footnotes, "footnote", "separator", "-1", generate_hex_id
    )
    add_note_separator(
        footnotes, "footnote", "continuationSeparator", "0", generate_hex_id
    )
    return etree.tostring(
        footnotes, xml_declaration=True, encoding="UTF-8", pretty_print=True
    )


def create_endnotes(generate_hex_id) -> bytes:
    """Create word/endnotes.xml."""
    w_ns = NAMESPACES["w"]
    w14_ns = NAMESPACES["w14"]
    endnotes = etree.Element(
        f"{{{w_ns}}}endnotes",
        nsmap={"w": w_ns, "w14": w14_ns},
    )
    add_note_separator(
        endnotes, "endnote", "separator", "-1", generate_hex_id
    )
    add_note_separator(
        endnotes, "endnote", "continuationSeparator", "0", generate_hex_id
    )
    return etree.tostring(
        endnotes, xml_declaration=True, encoding="UTF-8", pretty_print=True
    )


def add_note_separator(
    parent: XMLElement, tag: str, sep_tag: str, note_id: str, generate_hex_id
) -> None:
    """Add a note separator entry for footnotes or endnotes."""
    w_ns = NAMESPACES["w"]
    w14_ns = NAMESPACES["w14"]
    note = etree.SubElement(
        parent,
        f"{{{w_ns}}}{tag}",
        {f"{{{w_ns}}}type": sep_tag, f"{{{w_ns}}}id": note_id},
    )
    para = etree.SubElement(note, f"{{{w_ns}}}p")
    para.set(f"{{{w14_ns}}}paraId", generate_hex_id(8))
    para.set(f"{{{w14_ns}}}textId", "77777777")
    p_pr = etree.SubElement(para, f"{{{w_ns}}}pPr")
    spacing = etree.SubElement(p_pr, f"{{{w_ns}}}spacing")
    spacing.set(f"{{{w_ns}}}after", "0")
    spacing.set(f"{{{w_ns}}}line", "240")
    spacing.set(f"{{{w_ns}}}lineRule", "auto")
    run = etree.SubElement(para, f"{{{w_ns}}}r")
    etree.SubElement(run, f"{{{w_ns}}}{sep_tag}")


def create_font_table() -> bytes:
    """Create word/fontTable.xml."""
    return FONT_TABLE_XML.encode("utf-8")


def create_theme() -> bytes:
    """Create word/theme/theme1.xml."""
    return base64.b64decode(THEME_XML_B64)


def create_core_properties(
    title: str,
    author: str,
    timestamp: dt.datetime | None = None,
) -> bytes:
    """Create docProps/core.xml."""
    ts = timestamp or dt.datetime.now(dt.UTC)
    now = ts.strftime("%Y-%m-%dT%H:%M:%SZ")
    cp_ns = (
        "http://schemas.openxmlformats.org/package/"
        "2006/metadata/core-properties"
    )
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<cp:coreProperties xmlns:cp="{cp_ns}" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{title}</dc:title>"
        "<dc:subject></dc:subject>"
        f"<dc:creator>{author}</dc:creator>"
        "<cp:keywords></cp:keywords>"
        "<dc:description></dc:description>"
        "<cp:lastModifiedBy></cp:lastModifiedBy>"
        "<cp:revision>1</cp:revision>"
        "<dcterms:created "
        f'xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        "<dcterms:modified "
        f'xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )
    return core_xml.encode("utf-8")


def create_app_properties() -> bytes:
    """Create docProps/app.xml."""
    return APP_PROPERTIES_XML.encode("utf-8")
