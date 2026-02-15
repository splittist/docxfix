"""Integration tests for comment generation against golden corpus."""

import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pytest
from lxml import etree

from docxfix.generator import DocumentGenerator
from docxfix.spec import Comment, CommentReply, DocumentSpec
from docxfix.validator import validate_docx


def test_generate_comment_thread_fixture():
    """Test generating a fixture similar to comment-thread.docx golden."""
    spec = DocumentSpec(seed=42)
    
    # Create a comment with reply similar to the golden fixture
    reply = CommentReply(
        text="A reply comment.",
        author="Author",
        date=datetime(2026, 2, 14, 12, 0, 0),
    )
    
    comment = Comment(
        text="A comment.",
        anchor_text="dolor sit amet",
        author="Author",
        date=datetime(2026, 2, 14, 12, 0, 0),
        replies=[reply],
        resolved=False,
    )
    
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        comments=[comment]
    )
    
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "comment-thread-generated.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        
        # Validate the generated document (raises exception on error)
        validate_docx(output_path)
        
        with zipfile.ZipFile(output_path, "r") as docx_zip:
            # Verify structure matches expectations from golden fixture
            
            # 1. Check document.xml has comment markers
            doc_xml = docx_zip.read("word/document.xml")
            doc_root = etree.fromstring(doc_xml)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            
            # Should have 2 commentRangeStart (main comment + reply)
            comment_starts = doc_root.findall(".//w:commentRangeStart", namespaces=ns)
            assert len(comment_starts) == 2
            
            # Should have 2 commentRangeEnd
            comment_ends = doc_root.findall(".//w:commentRangeEnd", namespaces=ns)
            assert len(comment_ends) == 2
            
            # Should have 2 commentReference
            comment_refs = doc_root.findall(".//w:commentReference", namespaces=ns)
            assert len(comment_refs) == 2
            
            # 2. Check comments.xml has 2 comment elements
            comments_xml = docx_zip.read("word/comments.xml")
            comments_root = etree.fromstring(comments_xml)
            comments = comments_root.findall(".//w:comment", namespaces=ns)
            assert len(comments) == 2
            
            # Each comment should have id, author, initials
            for c in comments:
                assert c.get(f"{{{ns['w']}}}id") is not None
                assert c.get(f"{{{ns['w']}}}author") == "Author"
                assert c.get(f"{{{ns['w']}}}initials") is not None
            
            # 3. Check commentsExtended.xml structure
            comments_ext_xml = docx_zip.read("word/commentsExtended.xml")
            ns15 = {"w15": "http://schemas.microsoft.com/office/word/2012/wordml"}
            ext_root = etree.fromstring(comments_ext_xml)
            comment_exs = ext_root.findall(".//w15:commentEx", namespaces=ns15)
            assert len(comment_exs) == 2
            
            # Both should have paraId and done attributes
            for ex in comment_exs:
                assert ex.get(f"{{{ns15['w15']}}}paraId") is not None
                assert ex.get(f"{{{ns15['w15']}}}done") == "0"  # Not resolved
            
            # Second comment should have paraIdParent
            parent_refs = [
                ex.get(f"{{{ns15['w15']}}}paraIdParent")
                for ex in comment_exs
                if f"{{{ns15['w15']}}}paraIdParent" in ex.attrib
            ]
            assert len(parent_refs) == 1  # Only the reply has a parent

            # 4. Check commentsIds.xml exists and has correct structure
            assert "word/commentsIds.xml" in docx_zip.namelist()
            comments_ids_xml = docx_zip.read("word/commentsIds.xml")
            ns16cid = {"w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid"}
            ids_root = etree.fromstring(comments_ids_xml)
            comment_id_elems = ids_root.findall(".//w16cid:commentId", namespaces=ns16cid)
            assert len(comment_id_elems) == 2  # One for main comment, one for reply


def test_generate_resolved_comment_fixture():
    """Test generating a fixture with a resolved comment."""
    spec = DocumentSpec(seed=42)
    
    comment = Comment(
        text="This should be resolved.",
        anchor_text="dolor sit amet",
        author="Author",
        date=datetime(2026, 2, 14, 12, 0, 0),
        resolved=True,
    )
    
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        comments=[comment]
    )
    
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "resolved-comment-generated.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        
        # Validate the generated document (raises exception on error)
        validate_docx(output_path)
        
        with zipfile.ZipFile(output_path, "r") as docx_zip:
            # Check commentsExtended.xml has done="1"
            comments_ext_xml = docx_zip.read("word/commentsExtended.xml")
            ns15 = {"w15": "http://schemas.microsoft.com/office/word/2012/wordml"}
            ext_root = etree.fromstring(comments_ext_xml)
            comment_exs = ext_root.findall(".//w15:commentEx", namespaces=ns15)
            assert len(comment_exs) == 1
            
            done_attr = comment_exs[0].get(f"{{{ns15['w15']}}}done")
            assert done_attr == "1"


def test_generate_multiple_reply_levels():
    """Test generating a comment with multiple replies."""
    spec = DocumentSpec(seed=42)
    
    replies = [
        CommentReply(
            text="First reply.",
            author="User A",
            date=datetime(2026, 2, 14, 12, 0, 0),
        ),
        CommentReply(
            text="Second reply.",
            author="User B",
            date=datetime(2026, 2, 14, 12, 5, 0),
        ),
        CommentReply(
            text="Third reply.",
            author="User C",
            date=datetime(2026, 2, 14, 12, 10, 0),
        ),
    ]
    
    comment = Comment(
        text="Original comment.",
        anchor_text="dolor sit amet",
        author="Original Author",
        date=datetime(2026, 2, 14, 11, 0, 0),
        replies=replies,
    )
    
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.",
        comments=[comment]
    )
    
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "multi-reply-generated.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        
        # Validate the generated document (raises exception on error)
        validate_docx(output_path)
        
        with zipfile.ZipFile(output_path, "r") as docx_zip:
            # Should have 4 total comments (1 original + 3 replies)
            comments_xml = docx_zip.read("word/comments.xml")
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            comments_root = etree.fromstring(comments_xml)
            comments = comments_root.findall(".//w:comment", namespaces=ns)
            assert len(comments) == 4
            
            # Check commentsExtended has 3 replies with parent references
            comments_ext_xml = docx_zip.read("word/commentsExtended.xml")
            ns15 = {"w15": "http://schemas.microsoft.com/office/word/2012/wordml"}
            ext_root = etree.fromstring(comments_ext_xml)
            comment_exs = ext_root.findall(".//w15:commentEx", namespaces=ns15)
            assert len(comment_exs) == 4
            
            # 3 should have parent references
            parent_refs = [
                ex.get(f"{{{ns15['w15']}}}paraIdParent")
                for ex in comment_exs
                if f"{{{ns15['w15']}}}paraIdParent" in ex.attrib
            ]
            assert len(parent_refs) == 3
            
            # All replies should have the same parent
            assert len(set(parent_refs)) == 1


def test_comment_without_anchor_text_in_paragraph():
    """Test handling when anchor text is not found in paragraph."""
    spec = DocumentSpec(seed=42)
    
    comment = Comment(
        text="This is a comment.",
        anchor_text="nonexistent text",
        author="Author",
    )
    
    spec.add_paragraph(
        "Lorem ipsum dolor sit amet.",
        comments=[comment]
    )
    
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "comment-no-anchor.docx"
        generator = DocumentGenerator(spec)
        generator.generate(output_path)
        
        # Should still generate valid document (raises exception on error)
        validate_docx(output_path)
        
        with zipfile.ZipFile(output_path, "r") as docx_zip:
            # Should still have comment markers
            doc_xml = docx_zip.read("word/document.xml")
            doc_root = etree.fromstring(doc_xml)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            
            comment_starts = doc_root.findall(".//w:commentRangeStart", namespaces=ns)
            assert len(comment_starts) >= 1
