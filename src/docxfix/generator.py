"""Document generator for creating .docx files from specifications."""

import io
import random
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from docxfix.spec import ChangeType, Comment, DocumentSpec, Paragraph
from docxfix.xml_utils import XMLElement


class DocumentGenerator:
    """Generates .docx files from DocumentSpec."""

    # OOXML namespaces
    NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    }

    def __init__(self, spec: DocumentSpec) -> None:
        """Initialize generator with a document specification."""
        self.spec = spec
        self._revision_counter = 0
        self._comment_counter = 0
        self._comment_metadata = []  # Track comment metadata for multi-part generation
        
        # Initialize random seed if specified
        if spec.seed is not None:
            random.seed(spec.seed)

    def generate(self, output_path: str | Path) -> None:
        """
        Generate a .docx file at the specified path.

        Args:
            output_path: Path where the .docx file will be created
        """
        output_path = Path(output_path)
        
        # Check if document has comments
        has_comments = any(
            para.comments for para in self.spec.paragraphs
        )

        # Create a ZIP file (docx is a ZIP archive)
        with zipfile.ZipFile(
            output_path, "w", zipfile.ZIP_DEFLATED
        ) as docx_zip:
            # Add required files
            docx_zip.writestr("[Content_Types].xml", self._create_content_types(has_comments))
            docx_zip.writestr("_rels/.rels", self._create_rels())
            docx_zip.writestr(
                "word/_rels/document.xml.rels", self._create_document_rels(has_comments)
            )
            docx_zip.writestr("word/document.xml", self._create_document())
            
            # Add comment files if needed
            if has_comments:
                docx_zip.writestr("word/comments.xml", self._create_comments())
                docx_zip.writestr("word/commentsExtended.xml", self._create_comments_extended())
                docx_zip.writestr("word/commentsIds.xml", self._create_comments_ids())

    def _create_content_types(self, has_comments: bool = False) -> bytes:
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
        
        # Add comment content types if needed
        if has_comments:
            etree.SubElement(
                types,
                "Override",
                PartName="/word/comments.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.comments+xml"
                ),
            )
            etree.SubElement(
                types,
                "Override",
                PartName="/word/commentsExtended.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.commentsExtended+xml"
                ),
            )
            etree.SubElement(
                types,
                "Override",
                PartName="/word/commentsIds.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.commentsIds+xml"
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

    def _create_document_rels(self, has_comments: bool = False) -> bytes:
        """Create word/_rels/document.xml.rels."""
        rels = etree.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )
        
        # Add comment relationships if needed
        if has_comments:
            etree.SubElement(
                rels,
                "Relationship",
                Id="rId1",
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                Target="comments.xml",
            )
            etree.SubElement(
                rels,
                "Relationship",
                Id="rId2",
                Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                Target="commentsExtended.xml",
            )
            etree.SubElement(
                rels,
                "Relationship",
                Id="rId3",
                Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                Target="commentsIds.xml",
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
        w14_ns = self.NAMESPACES["w14"]
        para = etree.SubElement(body, f"{{{w_ns}}}p")
        
        # Generate unique paraId for paragraph
        para_id = self._generate_hex_id(8)
        para.set(f"{{{w14_ns}}}paraId", para_id)
        para.set(f"{{{w14_ns}}}textId", "77777777")  # Static textId for now

        # Handle different content types
        if para_spec.comments:
            self._add_paragraph_with_comments(para, para_spec)
        elif para_spec.tracked_changes:
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

    def _add_paragraph_with_comments(self, para: XMLElement, para_spec: Paragraph) -> None:
        """Add a paragraph with comment anchoring."""
        w_ns = self.NAMESPACES["w"]
        
        # Split text into before, anchor, and after parts
        # For simplicity, we'll find the anchor_text in the paragraph text
        # and add comment markers around it
        
        for comment in para_spec.comments:
            anchor_text = comment.anchor_text
            full_text = para_spec.text
            
            # Find anchor position
            if anchor_text not in full_text:
                # If anchor text not found, just comment the whole paragraph
                anchor_start = 0
                anchor_end = len(full_text)
                before_text = ""
                after_text = ""
            else:
                anchor_start = full_text.index(anchor_text)
                anchor_end = anchor_start + len(anchor_text)
                before_text = full_text[:anchor_start]
                after_text = full_text[anchor_end:]
            
            # Create comment ID and metadata
            comment_id = str(self._comment_counter)
            parent_para_id = self._generate_hex_id(8).upper()
            durable_id = self._generate_hex_id(8).upper()
            
            # Store metadata for later use in comment files
            self._comment_metadata.append({
                "id": comment_id,
                "para_id": parent_para_id,
                "durable_id": durable_id,
                "author": comment.author,
                "date": comment.date,
                "text": comment.text,
                "resolved": comment.resolved,
                "parent_para_id": None,  # No parent for main comment
            })
            
            self._comment_counter += 1
            
            # Handle replies
            reply_ids = []
            for reply in comment.replies:
                reply_id = str(self._comment_counter)
                reply_para_id = self._generate_hex_id(8).upper()
                reply_durable_id = self._generate_hex_id(8).upper()
                
                self._comment_metadata.append({
                    "id": reply_id,
                    "para_id": reply_para_id,
                    "durable_id": reply_durable_id,
                    "author": reply.author,
                    "date": reply.date,
                    "text": reply.text,
                    "resolved": comment.resolved,
                    "parent_para_id": parent_para_id,  # Link to parent comment
                })
                
                reply_ids.append(reply_id)
                self._comment_counter += 1
            
            # Add text before anchor
            if before_text:
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
                text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text_elem.text = before_text
            
            # Add comment range start for main comment
            etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": comment_id})
            
            # Add comment range starts for replies
            for reply_id in reply_ids:
                etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": reply_id})
            
            # Add the anchored text
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = anchor_text
            
            # Add comment range end for main comment
            etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": comment_id})
            
            # Add comment reference for main comment
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            etree.SubElement(run, f"{{{w_ns}}}commentReference", {f"{{{w_ns}}}id": comment_id})
            
            # Add comment range ends and references for replies
            for reply_id in reply_ids:
                etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": reply_id})
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                etree.SubElement(run, f"{{{w_ns}}}commentReference", {f"{{{w_ns}}}id": reply_id})
            
            # Add text after anchor
            if after_text:
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
                text_elem.text = after_text

    def _generate_hex_id(self, length: int = 8) -> str:
        """Generate a random hexadecimal ID of specified length."""
        return "".join(random.choices("0123456789ABCDEF", k=length))

    def _create_comments(self) -> bytes:
        """Create word/comments.xml."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        
        comments = etree.Element(
            f"{{{w_ns}}}comments",
            nsmap={
                "w": w_ns,
                "w14": w14_ns,
                "mc": self.NAMESPACES["mc"],
            },
        )
        comments.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14")
        
        # Add each comment
        for metadata in self._comment_metadata:
            comment = etree.SubElement(
                comments,
                f"{{{w_ns}}}comment",
                {
                    f"{{{w_ns}}}id": metadata["id"],
                    f"{{{w_ns}}}author": metadata["author"],
                    f"{{{w_ns}}}initials": metadata["author"][0] if metadata["author"] else "A",
                },
            )
            
            # Add comment paragraph
            para = etree.SubElement(comment, f"{{{w_ns}}}p")
            para.set(f"{{{w14_ns}}}paraId", metadata["para_id"])
            para.set(f"{{{w14_ns}}}textId", "77777777")
            
            # Add annotation reference run
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            etree.SubElement(run, f"{{{w_ns}}}annotationRef")
            
            # Add comment text
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = metadata["text"]
        
        return etree.tostring(
            comments, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_comments_extended(self) -> bytes:
        """Create word/commentsExtended.xml."""
        w15_ns = self.NAMESPACES["w15"]
        
        comments_ex = etree.Element(
            f"{{{w15_ns}}}commentsEx",
            nsmap={
                "w15": w15_ns,
                "mc": self.NAMESPACES["mc"],
            },
        )
        comments_ex.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w15")
        
        # Add each comment extension
        for metadata in self._comment_metadata:
            comment_ex = etree.SubElement(
                comments_ex,
                f"{{{w15_ns}}}commentEx",
                {
                    f"{{{w15_ns}}}paraId": metadata["para_id"],
                    f"{{{w15_ns}}}done": "1" if metadata["resolved"] else "0",
                },
            )
            
            # Add parent reference for replies
            if metadata["parent_para_id"]:
                comment_ex.set(f"{{{w15_ns}}}paraIdParent", metadata["parent_para_id"])
        
        return etree.tostring(
            comments_ex, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_comments_ids(self) -> bytes:
        """Create word/commentsIds.xml."""
        w16cid_ns = self.NAMESPACES["w16cid"]
        
        comments_ids = etree.Element(
            f"{{{w16cid_ns}}}commentsIds",
            nsmap={
                "w16cid": w16cid_ns,
                "mc": self.NAMESPACES["mc"],
            },
        )
        comments_ids.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w16cid")
        
        # Add each comment ID
        for metadata in self._comment_metadata:
            etree.SubElement(
                comments_ids,
                f"{{{w16cid_ns}}}commentId",
                {
                    f"{{{w16cid_ns}}}paraId": metadata["para_id"],
                    f"{{{w16cid_ns}}}durableId": metadata["durable_id"],
                },
            )
        
        return etree.tostring(
            comments_ids, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )
