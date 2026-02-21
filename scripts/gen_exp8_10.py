"""
Generate experiments exp8–exp10 using the ACTUAL docxfix generator.

exp4/5/6 (from gen_exp4_7.py) all failed — they used a custom minimal script
missing webSettings, footnotes, endnotes, fontTable, theme, core/app props.
exp3 (docxfix-generated + commentsIds added) partially worked (Case 2 failed).

These experiments use the full docxfix generator output as a base, then
post-process the ZIP to inject unique docId and/or commentsIds.xml.

  exp8: Full docxfix + unique docId + commentsIds
        Tests whether docxfix output + both fixes reliably threads all cases
  exp9: Full docxfix + unique docId + NO commentsIds
        Tests whether unique docId alone is sufficient with full docxfix structure
  exp10: Full docxfix + SHARED docId + commentsIds
         Tests whether unique docId is actually necessary (since v4 worked shared+no ids)

All use the same 4-case comment structure as exp3.
"""
import io
import random
import string
import zipfile
from pathlib import Path

from lxml import etree

SCRATCH = Path(__file__).parent.parent / "scratch_out"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

IGNORABLE = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"

WORD_NSMAP = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "mc": MC_NS,
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": W_NS,
    "w14": W14_NS,
    "w15": W15_NS,
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": W16CID_NS,
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

SHARED_DOC_ID = "4D55359B"  # The hardcoded ID from SETTINGS_XML constant

COMMENTS_IDS_CT = (
    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.commentsIds+xml"
)
COMMENTS_IDS_REL_TYPE = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)


def rand_hex(rng, n=8):
    return "".join(rng.choice(string.hexdigits.upper()[:16]) for _ in range(n))


def make_comments_ids_bytes(metadata):
    """Build commentsIds.xml from comment_metadata list."""
    root = etree.Element(
        f"{{{W16CID_NS}}}commentsIds", nsmap=WORD_NSMAP
    )
    root.set(f"{{{MC_NS}}}Ignorable", IGNORABLE)
    for m in metadata:
        etree.SubElement(
            root,
            f"{{{W16CID_NS}}}commentId",
            {
                f"{{{W16CID_NS}}}paraId": m["para_id"],
                f"{{{W16CID_NS}}}durableId": m["durable_id"],
            },
        )
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", pretty_print=True
    )


def inject_unique_doc_id(settings_bytes, new_doc_id):
    """Replace w14:docId value in settings.xml bytes."""
    text = settings_bytes.decode("utf-8")
    import re
    text = re.sub(
        r'w14:docId\s+w14:val="[0-9A-Fa-f]+"',
        f'w14:docId w14:val="{new_doc_id}"',
        text,
    )
    return text.encode("utf-8")


def add_comments_ids_to_zip(
    src_path, dst_path, new_doc_id, include_ids
):
    """
    Post-process a docxfix-generated .docx:
    1. Replace settings.xml with one having new_doc_id
    2. Optionally add commentsIds.xml and update content types + rels
    """
    # Read source ZIP
    src_entries = {}
    with zipfile.ZipFile(src_path, "r") as zf:
        for name in zf.namelist():
            src_entries[name] = zf.read(name)

    # Determine comment metadata from commentsExtended.xml
    metadata = []
    if "word/commentsExtended.xml" in src_entries:
        root = etree.fromstring(src_entries["word/commentsExtended.xml"])
        ns = {"w15": W15_NS}
        for ce in root.findall(f"{{{W15_NS}}}commentEx"):
            para_id = ce.get(f"{{{W15_NS}}}paraId")
            parent_para_id = ce.get(f"{{{W15_NS}}}paraIdParent")
            # durableId = paraId for parents (same as current code),
            # unique for replies — this matches what exp3 had
            if parent_para_id is None:
                durable_id = para_id  # parent: durableId = paraId (existing behavior)
            else:
                rng = random.Random(para_id)  # deterministic from para_id
                durable_id = rand_hex(rng, 8)
            metadata.append({
                "para_id": para_id,
                "durable_id": durable_id,
                "parent_para_id": parent_para_id,
            })

    # Update settings.xml
    if "word/settings.xml" in src_entries:
        src_entries["word/settings.xml"] = inject_unique_doc_id(
            src_entries["word/settings.xml"], new_doc_id
        )

    # Add commentsIds.xml
    if include_ids:
        src_entries["word/commentsIds.xml"] = make_comments_ids_bytes(metadata)

        # Update content types
        ct_root = etree.fromstring(src_entries["[Content_Types].xml"])
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        # Remove existing commentsIds override if present
        for elem in ct_root.findall(f"{{{ct_ns}}}Override"):
            if "commentsIds" in (elem.get("PartName") or ""):
                ct_root.remove(elem)
        # Add correct one
        etree.SubElement(
            ct_root,
            f"{{{ct_ns}}}Override",
            {
                "PartName": "/word/commentsIds.xml",
                "ContentType": COMMENTS_IDS_CT,
            },
        )
        src_entries["[Content_Types].xml"] = etree.tostring(
            ct_root, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

        # Update document rels
        rels_xml = src_entries["word/_rels/document.xml.rels"]
        rels_root = etree.fromstring(rels_xml)
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        # Find next rId
        existing_ids = [
            e.get("Id") for e in rels_root.findall(f"{{{rels_ns}}}Relationship")
        ]
        next_n = len(existing_ids) + 1
        new_rid = f"rId{next_n}"
        while new_rid in existing_ids:
            next_n += 1
            new_rid = f"rId{next_n}"
        etree.SubElement(
            rels_root,
            f"{{{rels_ns}}}Relationship",
            {
                "Id": new_rid,
                "Type": COMMENTS_IDS_REL_TYPE,
                "Target": "commentsIds.xml",
            },
        )
        src_entries["word/_rels/document.xml.rels"] = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    # Write output ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in src_entries.items():
            zf.writestr(name, data)
    dst_path.write_bytes(buf.getvalue())
    print(f"  wrote {dst_path.name} ({len(src_entries)} entries)")


def make_base_docx(seed):
    """Generate a base docxfix document with the standard 4-case structure."""
    from docxfix.spec import Comment, CommentReply, DocumentSpec
    from docxfix.generator import DocumentGenerator

    spec = (
        DocumentSpec(title="Threading Test", author="Test", seed=seed)
        .add_paragraph("Case 1 (standalone comment):")
        .add_paragraph(
            "This is Case 1 with one standalone comment.",
            comments=[
                Comment(
                    text="Standalone comment",
                    anchor_text="one standalone comment",
                    author="Alice",
                )
            ],
        )
        .add_paragraph("Case 2 (parent + 1 reply):")
        .add_paragraph(
            "This is Case 2 with one comment and one reply.",
            comments=[
                Comment(
                    text="Parent comment",
                    anchor_text="one comment and one reply",
                    author="Alice",
                    replies=[
                        CommentReply(text="Reply to parent", author="Bob"),
                    ],
                )
            ],
        )
        .add_paragraph("Case 3 (parent + 2 replies):")
        .add_paragraph(
            "This is Case 3 with one comment and two replies.",
            comments=[
                Comment(
                    text="Parent comment",
                    anchor_text="one comment and two replies",
                    author="Alice",
                    replies=[
                        CommentReply(text="Reply 1", author="Bob"),
                        CommentReply(text="Reply 2", author="Carol"),
                    ],
                )
            ],
        )
        .add_paragraph("Case 4 (two independent):")
        .add_paragraph(
            "This is Case 4 with two independent comments.",
            comments=[
                Comment(
                    text="Independent comment A",
                    anchor_text="two independent comments",
                    author="Alice",
                ),
                Comment(
                    text="Independent comment B",
                    anchor_text="two independent comments",
                    author="Bob",
                ),
            ],
        )
    )

    tmp = SCRATCH / "_tmp_base.docx"
    DocumentGenerator(spec).generate(tmp)
    return tmp


def main():
    SCRATCH.mkdir(exist_ok=True)

    print("Generating base docxfix document...")
    base = make_base_docx(seed=42)
    print(f"  base: {base.name} ({base.stat().st_size} bytes)")

    rng = random.Random(55)

    print("\nGenerating exp8–exp10...")

    # exp8: unique docId + commentsIds
    add_comments_ids_to_zip(
        base,
        SCRATCH / "exp8-docxfix-unique-docid-with-ids.docx",
        new_doc_id=rand_hex(rng),
        include_ids=True,
    )

    # exp9: unique docId + no commentsIds
    add_comments_ids_to_zip(
        base,
        SCRATCH / "exp9-docxfix-unique-docid-no-ids.docx",
        new_doc_id=rand_hex(rng),
        include_ids=False,
    )

    # exp10: shared docId + commentsIds
    add_comments_ids_to_zip(
        base,
        SCRATCH / "exp10-docxfix-shared-docid-with-ids.docx",
        new_doc_id=SHARED_DOC_ID,
        include_ids=True,
    )

    # Clean up temp file
    base.unlink()

    print("\nExperiments generated:")
    print("  exp8-docxfix-unique-docid-with-ids.docx")
    print("    → Full docxfix + unique docId + commentsIds (both fixes)")
    print("    → If works: the fix is confirmed; minimal script was missing files")
    print("    → If Case 2 fails: durableId=paraId is still the issue")
    print()
    print("  exp9-docxfix-unique-docid-no-ids.docx")
    print("    → Full docxfix + unique docId + NO commentsIds")
    print("    → If works: unique docId alone is sufficient with full structure")
    print("    → If fails: commentsIds is also needed")
    print()
    print("  exp10-docxfix-shared-docid-with-ids.docx")
    print("    → Full docxfix + SHARED docId (4D55359B) + commentsIds")
    print("    → If works: docId uniqueness is not needed (commentsIds alone fixes)")
    print("    → If fails: unique docId is required")


if __name__ == "__main__":
    main()
