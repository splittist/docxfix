"""Integration tests for numbering generation against golden corpus."""

import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec, NumberedParagraph


def test_generate_legal_list_fixture(tmp_path):
    """Test generating a fixture matching legal-list.docx from the corpus."""
    # Based on corpus/legal-list.md:
    # Five paragraphs with numbering levels 0, 1, 0, 1, 2
    # Should produce: 1., 1.1., 2., 2.1., 2.1.1.
    
    spec = DocumentSpec()
    spec.add_paragraph(text="Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    spec.add_paragraph(text="Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.",
            numbering=NumberedParagraph(level=1, numbering_id=1),)
    spec.add_paragraph(text="Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    spec.add_paragraph(text="Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.",
            numbering=NumberedParagraph(level=1, numbering_id=1),)
    spec.add_paragraph(text="Proin pharetra nonummy pede. Mauris et orci.",
            numbering=NumberedParagraph(level=2, numbering_id=1),)
    
    output_path = tmp_path / "legal-list-test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify structure matches golden corpus expectations
    with zipfile.ZipFile(output_path) as z:
        # Check document.xml
        doc_content = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        # Should have 5 paragraphs
        paras = doc_root.findall(".//w:p", namespaces)
        assert len(paras) == 5
        
        # All paragraphs should have pPr with pStyle="ListParagraph"
        for para in paras:
            pPr = para.find("w:pPr", namespaces)
            assert pPr is not None
            
            pStyle = pPr.find("w:pStyle", namespaces)
            assert pStyle is not None
            assert pStyle.get(f"{{{namespaces['w']}}}val") == "ListParagraph"
        
        # All paragraphs should have numPr with numId="1"
        for para in paras:
            numPr = para.find(".//w:numPr", namespaces)
            assert numPr is not None
            
            numId = numPr.find("w:numId", namespaces)
            assert numId is not None
            assert numId.get(f"{{{namespaces['w']}}}val") == "1"
        
        # Check ilvl values: 0, 1, 0, 1, 2
        expected_levels = ["0", "1", "0", "1", "2"]
        for para, expected_level in zip(paras, expected_levels):
            numPr = para.find(".//w:numPr", namespaces)
            ilvl = numPr.find("w:ilvl", namespaces)
            assert ilvl is not None
            assert ilvl.get(f"{{{namespaces['w']}}}val") == expected_level
        
        # Check numbering.xml structure
        num_content = z.read("word/numbering.xml")
        num_root = etree.fromstring(num_content)
        
        # Should have one abstractNum with abstractNumId="0"
        abstractNums = num_root.findall("w:abstractNum", namespaces)
        assert len(abstractNums) == 1
        assert abstractNums[0].get(f"{{{namespaces['w']}}}abstractNumId") == "0"
        
        # Should have multiLevelType="multilevel"
        multiLevelType = abstractNums[0].find("w:multiLevelType", namespaces)
        assert multiLevelType is not None
        assert multiLevelType.get(f"{{{namespaces['w']}}}val") == "multilevel"
        
        # Should have one num with numId="1" pointing to abstractNumId="0"
        nums = num_root.findall("w:num", namespaces)
        assert len(nums) == 1
        assert nums[0].get(f"{{{namespaces['w']}}}numId") == "1"
        
        abstractNumId = nums[0].find("w:abstractNumId", namespaces)
        assert abstractNumId is not None
        assert abstractNumId.get(f"{{{namespaces['w']}}}val") == "0"


def test_legal_numbering_format_patterns(tmp_path):
    """Test that legal numbering patterns are correct for each level."""
    spec = DocumentSpec()
    
    # Add paragraphs at different levels to test all level formats
    for level in range(9):
        spec.add_paragraph(text=f"Level {level} item",
                numbering=NumberedParagraph(level=level, numbering_id=1),)
    
    output_path = tmp_path / "all-levels.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify numbering.xml has correct level text patterns
    with zipfile.ZipFile(output_path) as z:
        num_content = z.read("word/numbering.xml")
        num_root = etree.fromstring(num_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        abstractNum = num_root.find("w:abstractNum", namespaces)
        lvls = abstractNum.findall("w:lvl", namespaces)
        
        # Expected level text patterns for legal-style numbering
        expected_patterns = [
            "%1.",
            "%1.%2.",
            "%1.%2.%3.",
            "%1.%2.%3.%4.",
            "%1.%2.%3.%4.%5.",
            "%1.%2.%3.%4.%5.%6.",
            "%1.%2.%3.%4.%5.%6.%7.",
            "%1.%2.%3.%4.%5.%6.%7.%8.",
            "%1.%2.%3.%4.%5.%6.%7.%8.%9.",
        ]
        
        assert len(lvls) == 9
        
        for lvl, expected_pattern in zip(lvls, expected_patterns):
            lvlText = lvl.find("w:lvlText", namespaces)
            assert lvlText is not None
            assert lvlText.get(f"{{{namespaces['w']}}}val") == expected_pattern
            
            # All should be decimal format
            numFmt = lvl.find("w:numFmt", namespaces)
            assert numFmt is not None
            assert numFmt.get(f"{{{namespaces['w']}}}val") == "decimal"


def test_numbering_indentation_values(tmp_path):
    """Test that indentation values are correctly set for each level."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Test",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Expected indentation values (left, hanging) for legal numbering
    expected_indents = [
        (360, 360),
        (792, 432),
        (1224, 504),
        (1728, 648),
        (2232, 792),
        (2736, 936),
        (3240, 1080),
        (3744, 1224),
        (4320, 1440),
    ]
    
    with zipfile.ZipFile(output_path) as z:
        num_content = z.read("word/numbering.xml")
        num_root = etree.fromstring(num_content)
        
        namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        abstractNum = num_root.find("w:abstractNum", namespaces)
        lvls = abstractNum.findall("w:lvl", namespaces)
        
        for lvl, (expected_left, expected_hanging) in zip(lvls, expected_indents):
            # Find pPr > ind element
            pPr = lvl.find("w:pPr", namespaces)
            assert pPr is not None
            
            ind = pPr.find("w:ind", namespaces)
            assert ind is not None
            
            left = ind.get(f"{{{namespaces['w']}}}left")
            hanging = ind.get(f"{{{namespaces['w']}}}hanging")
            
            level_num = lvl.get(f"{{{namespaces['w']}}}ilvl")
            assert left == str(expected_left), f"Level {level_num}: expected left={expected_left}, got {left}"
            assert hanging == str(expected_hanging), f"Level {level_num}: expected hanging={expected_hanging}, got {hanging}"


def test_numbering_without_comments(tmp_path):
    """Test that numbering works correctly when no comments are present."""
    spec = DocumentSpec()
    spec.add_paragraph(text="Numbered item",
            numbering=NumberedParagraph(level=0, numbering_id=1),)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify relationship IDs are correct without comments
    with zipfile.ZipFile(output_path) as z:
        rels_content = z.read("word/_rels/document.xml.rels")
        rels_root = etree.fromstring(rels_content)
        
        rels = rels_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        
        # When no comments, numbering should use rId1 and styles rId2
        numbering_rel = None
        styles_rel = None
        
        for rel in rels:
            if "numbering" in rel.get("Type"):
                numbering_rel = rel
            elif "styles" in rel.get("Type"):
                styles_rel = rel
        
        assert numbering_rel is not None
        assert numbering_rel.get("Id") == "rId1"
        
        assert styles_rel is not None
        assert styles_rel.get("Id") == "rId2"


def test_numbering_with_comments(tmp_path):
    """Test that numbering and comments work together correctly."""
    from docxfix.spec import Comment
    
    spec = DocumentSpec()
    spec.add_paragraph(text="Numbered item with comment",
            numbering=NumberedParagraph(level=0, numbering_id=1),
            comments=[Comment(text="Test comment", anchor_text="comment")],)
    
    output_path = tmp_path / "test.docx"
    generator = DocumentGenerator(spec)
    generator.generate(output_path)
    
    # Verify all required files are present
    with zipfile.ZipFile(output_path) as z:
        assert "word/numbering.xml" in z.namelist()
        assert "word/styles.xml" in z.namelist()
        assert "word/comments.xml" in z.namelist()
        assert "word/commentsExtended.xml" in z.namelist()
        assert "word/commentsIds.xml" in z.namelist()
        
        # Verify relationship IDs are correct with comments
        rels_content = z.read("word/_rels/document.xml.rels")
        rels_root = etree.fromstring(rels_content)
        
        rels = rels_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        
        # When comments present, numbering should use rId4 and styles rId5
        numbering_rel = None
        styles_rel = None
        
        for rel in rels:
            if "numbering" in rel.get("Type"):
                numbering_rel = rel
            elif "styles" in rel.get("Type"):
                styles_rel = rel
        
        assert numbering_rel is not None
        assert numbering_rel.get("Id") == "rId4"
        
        assert styles_rel is not None
        assert styles_rel.get("Id") == "rId5"
