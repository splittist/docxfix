"""Tests for numbering generation in DocumentGenerator."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec, NumberedParagraph


def test_generator_simple_numbered_paragraph(tmp_path):
    """Test generating a document with a single numbered paragraph."""
    spec = DocumentSpec()
    spec.add_paragraph(
        text="First numbered item",
        numbering=NumberedParagraph(level=0, numbering_id=1),
    )
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify the file was created
    assert output_path.exists()
    
    # Verify it's a valid ZIP
    with zipfile.ZipFile(output_path) as z:
        # Check numbering.xml exists
        assert "word/numbering.xml" in z.namelist()
        
        # Check styles.xml exists
        assert "word/styles.xml" in z.namelist()
        
        # Parse document.xml
        doc_content = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_content)
        
        # Find the paragraph with numbering properties
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paras = doc_root.findall(".//w:p", namespaces)
        assert len(paras) == 1
        
        # Check for pPr with numPr
        pPr = paras[0].find("w:pPr", namespaces)
        assert pPr is not None
        
        # Check for pStyle (ListParagraph)
        pStyle = pPr.find("w:pStyle", namespaces)
        assert pStyle is not None
        assert pStyle.get(f"{{{namespaces['w']}}}val") == "ListParagraph"
        
        # Check for numPr
        numPr = pPr.find("w:numPr", namespaces)
        assert numPr is not None
        
        # Check ilvl (indentation level)
        ilvl = numPr.find("w:ilvl", namespaces)
        assert ilvl is not None
        assert ilvl.get(f"{{{namespaces['w']}}}val") == "0"
        
        # Check numId
        numId = numPr.find("w:numId", namespaces)
        assert numId is not None
        assert numId.get(f"{{{namespaces['w']}}}val") == "1"


def test_generator_multilevel_numbering(tmp_path):
    """Test generating a document with multilevel numbering."""
    spec = DocumentSpec()
    spec.add_paragraph(text="First level item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    spec.add_paragraph(text="Second level item",
            numbering=NumberedParagraph(level=1, numbering_id=1),)
    spec.add_paragraph(text="Another first level item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    spec.add_paragraph(text="Another second level item",
            numbering=NumberedParagraph(level=1, numbering_id=1),)
    spec.add_paragraph(text="Third level item",
            numbering=NumberedParagraph(level=2, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify the file was created
    assert output_path.exists()
    
    # Parse document.xml
    with zipfile.ZipFile(output_path) as z:
        doc_content = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paras = doc_root.findall(".//w:p", namespaces)
        assert len(paras) == 5
        
        # Check levels: 0, 1, 0, 1, 2
        expected_levels = ["0", "1", "0", "1", "2"]
        for para, expected_level in zip(paras, expected_levels):
            numPr = para.find(".//w:numPr", namespaces)
            assert numPr is not None
            
            ilvl = numPr.find("w:ilvl", namespaces)
            assert ilvl is not None
            assert ilvl.get(f"{{{namespaces['w']}}}val") == expected_level


def test_numbering_xml_structure(tmp_path):
    """Test that numbering.xml has the correct structure."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Test item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Parse numbering.xml
    with zipfile.ZipFile(output_path) as z:
        num_content = z.read("word/numbering.xml")
        num_root = etree.fromstring(num_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        # Check for abstractNum element
        abstractNums = num_root.findall("w:abstractNum", namespaces)
        assert len(abstractNums) == 1
        
        abstractNum = abstractNums[0]
        assert abstractNum.get(f"{{{namespaces['w']}}}abstractNumId") == "0"
        
        # Check for multiLevelType
        multiLevelType = abstractNum.find("w:multiLevelType", namespaces)
        assert multiLevelType is not None
        assert multiLevelType.get(f"{{{namespaces['w']}}}val") == "multilevel"
        
        # Check for lvl elements (should have 9 levels: 0-8)
        lvls = abstractNum.findall("w:lvl", namespaces)
        assert len(lvls) == 9
        
        # Check first level (0)
        lvl0 = lvls[0]
        assert lvl0.get(f"{{{namespaces['w']}}}ilvl") == "0"
        
        # Check level 0 properties
        start = lvl0.find("w:start", namespaces)
        assert start is not None
        assert start.get(f"{{{namespaces['w']}}}val") == "1"
        
        numFmt = lvl0.find("w:numFmt", namespaces)
        assert numFmt is not None
        assert numFmt.get(f"{{{namespaces['w']}}}val") == "decimal"
        
        lvlText = lvl0.find("w:lvlText", namespaces)
        assert lvlText is not None
        assert lvlText.get(f"{{{namespaces['w']}}}val") == "%1."
        
        lvlJc = lvl0.find("w:lvlJc", namespaces)
        assert lvlJc is not None
        assert lvlJc.get(f"{{{namespaces['w']}}}val") == "left"
        
        # Check level 1 (should have "%1.%2." format)
        lvl1 = lvls[1]
        assert lvl1.get(f"{{{namespaces['w']}}}ilvl") == "1"
        lvlText1 = lvl1.find("w:lvlText", namespaces)
        assert lvlText1.get(f"{{{namespaces['w']}}}val") == "%1.%2."
        
        # Check level 2 (should have "%1.%2.%3." format)
        lvl2 = lvls[2]
        assert lvl2.get(f"{{{namespaces['w']}}}ilvl") == "2"
        lvlText2 = lvl2.find("w:lvlText", namespaces)
        assert lvlText2.get(f"{{{namespaces['w']}}}val") == "%1.%2.%3."
        
        # Check for concrete num element
        nums = num_root.findall("w:num", namespaces)
        assert len(nums) == 1
        
        num = nums[0]
        assert num.get(f"{{{namespaces['w']}}}numId") == "1"
        
        # Check abstractNumId reference
        abstractNumId = num.find("w:abstractNumId", namespaces)
        assert abstractNumId is not None
        assert abstractNumId.get(f"{{{namespaces['w']}}}val") == "0"


def test_styles_xml_structure(tmp_path):
    """Test that styles.xml has the correct structure."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Test item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Parse styles.xml
    with zipfile.ZipFile(output_path) as z:
        styles_content = z.read("word/styles.xml")
        styles_root = etree.fromstring(styles_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        # Check for style element with styleId="ListParagraph"
        styles = styles_root.findall("w:style", namespaces)
        assert len(styles) >= 1
        
        list_para_style = None
        for style in styles:
            if style.get(f"{{{namespaces['w']}}}styleId") == "ListParagraph":
                list_para_style = style
                break
        
        assert list_para_style is not None
        assert list_para_style.get(f"{{{namespaces['w']}}}type") == "paragraph"
        
        # Check for name
        name = list_para_style.find("w:name", namespaces)
        assert name is not None
        assert name.get(f"{{{namespaces['w']}}}val") == "List Paragraph"
        
        # Check for basedOn
        basedOn = list_para_style.find("w:basedOn", namespaces)
        assert basedOn is not None
        assert basedOn.get(f"{{{namespaces['w']}}}val") == "Normal"


def test_content_types_includes_numbering(tmp_path):
    """Test that [Content_Types].xml includes numbering and styles parts."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Test item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Parse [Content_Types].xml
    with zipfile.ZipFile(output_path) as z:
        ct_content = z.read("[Content_Types].xml")
        ct_root = etree.fromstring(ct_content)
        
        # Find Override elements
        overrides = ct_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
        )
        
        # Check for numbering.xml
        numbering_found = False
        styles_found = False
        
        for override in overrides:
            part_name = override.get("PartName")
            if part_name == "/word/numbering.xml":
                numbering_found = True
                assert "numbering" in override.get("ContentType")
            elif part_name == "/word/styles.xml":
                styles_found = True
                assert "styles" in override.get("ContentType")
        
        assert numbering_found, "numbering.xml not found in Content_Types"
        assert styles_found, "styles.xml not found in Content_Types"


def test_document_rels_includes_numbering(tmp_path):
    """Test that word/_rels/document.xml.rels includes numbering and styles relationships."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Test item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Parse document.xml.rels
    with zipfile.ZipFile(output_path) as z:
        rels_content = z.read("word/_rels/document.xml.rels")
        rels_root = etree.fromstring(rels_content)
        
        # Find Relationship elements
        rels = rels_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        
        # Check for numbering and styles relationships
        numbering_found = False
        styles_found = False
        
        for rel in rels:
            rel_type = rel.get("Type")
            target = rel.get("Target")
            
            if "numbering" in rel_type:
                numbering_found = True
                assert target == "numbering.xml"
            elif "styles" in rel_type:
                styles_found = True
                assert target == "styles.xml"
        
        assert numbering_found, "Numbering relationship not found"
        assert styles_found, "Styles relationship not found"


def test_mixed_numbered_and_plain_paragraphs(tmp_path):
    """Test document with both numbered and plain paragraphs."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Plain paragraph 1")
    spec.add_paragraph(text="Numbered item 1",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    spec.add_paragraph(text="Plain paragraph 2")
    spec.add_paragraph(text="Numbered item 2",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Parse document.xml
    with zipfile.ZipFile(output_path) as z:
        doc_content = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paras = doc_root.findall(".//w:p", namespaces)
        assert len(paras) == 4
        
        # First paragraph should not have numPr
        assert paras[0].find(".//w:numPr", namespaces) is None
        
        # Second paragraph should have numPr
        assert paras[1].find(".//w:numPr", namespaces) is not None
        
        # Third paragraph should not have numPr
        assert paras[2].find(".//w:numPr", namespaces) is None
        
        # Fourth paragraph should have numPr
        assert paras[3].find(".//w:numPr", namespaces) is not None
