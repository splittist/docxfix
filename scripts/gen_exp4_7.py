"""
Generate experiments exp4–exp7 to diagnose the Case 2 threading failure.

Context: exp3 (unique docId + commentsIds) partially works:
  - Case 1 (standalone) ✅
  - Case 2 (parent + 1 reply) ❌ NOT threaded
  - Case 3 (parent + 2 replies) ✅
  - Case 4 (2 independents) ✅

Hypotheses to test:
  exp4: Case 2 in isolation — is failure structural or context-dependent?
  exp5: Cases swapped — Case 3 structure first, then Case 2 structure
        Tests whether it's position-dependent or structure-dependent (1 vs 2 replies)
  exp6: Unique durableIds for ALL comments (parents too, not paraId=durableId)
        Tests durableId hypothesis
  exp7: 3 comments but thread of size 2 (parent + 1 reply + 1 standalone)
        Tests whether exactly 1 reply fails, or only when surrounded by other threads

All experiments use unique docId and include commentsIds.xml.
"""
import io
import random
import string
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

SCRATCH = Path(__file__).parent.parent / "scratch_out"

# Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

IGNORABLE = (
    "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"
)

WORD_NSMAP = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "mc": MC_NS,
    "o": "urn:schemas-microsoft-com:office:office",
    "r": R_NS,
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

DATE_STR = "2025-02-21T10:00:00Z"


def rand_hex(rng, n=8):
    return "".join(rng.choice(string.hexdigits.upper()[:16]) for _ in range(n))


def xml_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True)


def w(tag):
    return f"{{{W_NS}}}{tag}"

def w14(tag):
    return f"{{{W14_NS}}}{tag}"

def w15(tag):
    return f"{{{W15_NS}}}{tag}"

def w16cid(tag):
    return f"{{{W16CID_NS}}}{tag}"

def mc(tag):
    return f"{{{MC_NS}}}{tag}"


# ── document.xml ──────────────────────────────────────────────────────────────

def make_text_para(text, para_id, rng):
    """Paragraph with just text, no comments."""
    para = etree.Element(w("p"), nsmap=WORD_NSMAP)
    para.set(w14("paraId"), para_id)
    para.set(w14("textId"), rand_hex(rng))
    run = etree.SubElement(para, w("r"))
    t = etree.SubElement(run, w("t"))
    t.text = text
    return para


def make_comment_para(text, anchor, comment_ids, reply_ids, para_id, rng):
    """Paragraph with one anchored comment (+ replies)."""
    para = etree.Element(w("p"), nsmap=WORD_NSMAP)
    para.set(w14("paraId"), para_id)
    para.set(w14("textId"), rand_hex(rng))

    pre, _, post = text.partition(anchor)

    if pre:
        r = etree.SubElement(para, w("r"))
        t = etree.SubElement(r, w("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = pre

    # comment range starts
    for cid in comment_ids:
        etree.SubElement(para, w("commentRangeStart"), {w("id"): str(cid)})
    for rid in reply_ids:
        etree.SubElement(para, w("commentRangeStart"), {w("id"): str(rid)})

    r = etree.SubElement(para, w("r"))
    t = etree.SubElement(r, w("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = anchor

    for cid in comment_ids:
        etree.SubElement(para, w("commentRangeEnd"), {w("id"): str(cid)})
        rr = etree.SubElement(para, w("r"))
        _add_comment_ref(rr, str(cid))

    for rid in reply_ids:
        etree.SubElement(para, w("commentRangeEnd"), {w("id"): str(rid)})
        rr = etree.SubElement(para, w("r"))
        _add_comment_ref(rr, str(rid))

    if post:
        r = etree.SubElement(para, w("r"))
        t = etree.SubElement(r, w("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = post

    return para


def _add_comment_ref(run, cid):
    rpr = etree.SubElement(run, w("rPr"))
    rs = etree.SubElement(rpr, w("rStyle"))
    rs.set(w("val"), "CommentReference")
    etree.SubElement(run, w("commentReference"), {w("id"): cid})


# ── comments.xml ─────────────────────────────────────────────────────────────

def make_comment_elem(cid, para_id, author, text, rng):
    c = etree.Element(w("comment"), {
        w("id"): str(cid),
        w("author"): author,
        w("date"): DATE_STR,
        w("initials"): "".join(p[0].upper() for p in author.split()),
    })
    p = etree.SubElement(c, w("p"))
    p.set(w14("paraId"), para_id)
    p.set(w14("textId"), "77777777")
    ppr = etree.SubElement(p, w("pPr"))
    ps = etree.SubElement(ppr, w("pStyle"))
    ps.set(w("val"), "CommentText")
    run = etree.SubElement(p, w("r"))
    rpr = etree.SubElement(run, w("rPr"))
    rs = etree.SubElement(rpr, w("rStyle"))
    rs.set(w("val"), "CommentReference")
    etree.SubElement(run, w("annotationRef"))
    run2 = etree.SubElement(p, w("r"))
    t = etree.SubElement(run2, w("t"))
    t.text = text
    return c


# ── commentsExtended.xml ──────────────────────────────────────────────────────

def make_comments_extended(entries):
    """entries: list of (para_id, parent_para_id_or_None, resolved)"""
    root = etree.Element(w15("commentsEx"), nsmap=WORD_NSMAP)
    root.set(mc("Ignorable"), IGNORABLE)
    for (para_id, parent_para_id, resolved) in entries:
        attrs = {
            w15("paraId"): para_id,
            w15("done"): "1" if resolved else "0",
        }
        ce = etree.SubElement(root, w15("commentEx"), attrs)
        if parent_para_id:
            ce.set(w15("paraIdParent"), parent_para_id)
    return xml_bytes(root)


# ── commentsIds.xml ───────────────────────────────────────────────────────────

def make_comments_ids(entries):
    """entries: list of (para_id, durable_id)"""
    root = etree.Element(w16cid("commentsIds"), nsmap=WORD_NSMAP)
    root.set(mc("Ignorable"), IGNORABLE)
    for (para_id, durable_id) in entries:
        etree.SubElement(root, w16cid("commentId"), {
            w16cid("paraId"): para_id,
            w16cid("durableId"): durable_id,
        })
    return xml_bytes(root)


# ── settings.xml (unique docId) ───────────────────────────────────────────────

def make_settings(doc_id):
    return f"""<?xml version='1.0' encoding='UTF-8'?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:rsids>
    <w:rsidDel w:val="00000001"/>
  </w:rsids>
  <w14:docId w14:val="{doc_id}"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>""".encode()


# ── relationships ─────────────────────────────────────────────────────────────

RELS_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

DOC_RELS_TMPL = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsExtended" Target="commentsExtended.xml"/>
{extra_rels}
</Relationships>"""

COMMENTS_IDS_REL = '  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>'

ROOT_RELS = RELS_HEADER + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

CONTENT_TYPES_WITH_IDS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
  <Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats.officedocument.wordprocessingml.commentsIds+xml"/>
</Types>"""

CONTENT_TYPES_NO_IDS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
</Types>"""

MINIMAL_STYLES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="Comment Reference"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CommentText">
    <w:name w:val="Comment Text"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
</w:styles>"""


# ── document body builder ─────────────────────────────────────────────────────

def make_document(body_paras, rng):
    doc = etree.Element(w("document"), nsmap=WORD_NSMAP)
    body = etree.SubElement(doc, w("body"))
    for p in body_paras:
        body.append(p)
    # final sectPr
    etree.SubElement(body, w("sectPr"))
    return xml_bytes(doc)


def make_comments_xml(comment_elems):
    root = etree.Element(w("comments"), nsmap=WORD_NSMAP)
    root.set(mc("Ignorable"), IGNORABLE)
    for c in comment_elems:
        root.append(c)
    return xml_bytes(root)


# ── ZIP writer ────────────────────────────────────────────────────────────────

def write_docx(path, doc_bytes, comments_bytes, comments_ex_bytes,
               comments_ids_bytes, settings_bytes, include_ids=True):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",
                    CONTENT_TYPES_WITH_IDS if include_ids else CONTENT_TYPES_NO_IDS)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("word/document.xml", doc_bytes)
        zf.writestr("word/styles.xml", MINIMAL_STYLES)
        zf.writestr("word/settings.xml", settings_bytes)
        zf.writestr("word/comments.xml", comments_bytes)
        zf.writestr("word/commentsExtended.xml", comments_ex_bytes)
        if include_ids:
            zf.writestr("word/commentsIds.xml", comments_ids_bytes)
        extra = COMMENTS_IDS_REL if include_ids else ""
        zf.writestr("word/_rels/document.xml.rels",
                    RELS_HEADER + DOC_RELS_TMPL.format(extra_rels=extra))
    path.write_bytes(buf.getvalue())
    print(f"  wrote {path.name}")


# ── Experiment builders ───────────────────────────────────────────────────────

def build_exp4(rng):
    """
    exp4: ONLY Case 2 in isolation (unique docId + commentsIds).
    1 paragraph: parent comment + 1 reply.
    Tests whether Case 2 failure is structural or context-dependent.
    """
    doc_id = rand_hex(rng)
    cid = 0  # comment id counter

    # IDs
    p_para = rand_hex(rng)
    p_durable = rand_hex(rng)   # unique durable (same bug as exp3: we'll match)
    r_para = rand_hex(rng)
    r_durable = rand_hex(rng)

    # Body: intro + one comment para
    dp1 = rand_hex(rng)
    dp2 = rand_hex(rng)

    body_paras = [
        make_text_para("exp4: Case 2 isolation (parent + 1 reply)", dp1, rng),
        make_comment_para(
            "This paragraph has one comment with one reply.",
            "one comment with one reply",
            [0],   # parent cid
            [1],   # reply cid
            dp2, rng
        ),
    ]

    doc_bytes = make_document(body_paras, rng)

    parent_elem = make_comment_elem(0, p_para, "Alice Tester", "Parent comment", rng)
    reply_elem = make_comment_elem(1, r_para, "Bob Tester", "Reply to parent", rng)
    comments_bytes = make_comments_xml([parent_elem, reply_elem])

    # commentsExtended: reply has paraIdParent = parent's paraId
    extended_entries = [
        (p_para, None, False),
        (r_para, p_para, False),
    ]
    # commentsIds: parent durableId = p_para (same bug as exp3)
    ids_entries = [
        (p_para, p_para),   # durableId = paraId (bug)
        (r_para, r_durable),
    ]

    return (
        make_comments_extended(extended_entries),
        make_comments_ids(ids_entries),
        make_settings(doc_id),
        doc_bytes,
        comments_bytes,
    )


def build_exp5(rng):
    """
    exp5: Case 3 structure first, then Case 2 structure (swapped order vs exp3).
    Tests position vs structure dependency.
    Full 4 cases but Cases 2 and 3 swapped.
    """
    doc_id = rand_hex(rng)

    # IDs for cases — generate all upfront
    # [case1: standalone], [case2: parent+1reply], [case3: parent+2replies], [case4a, case4b]
    # Swapped: case3 comes first, then case2

    ids = {
        "c1_para": rand_hex(rng), "c1_dur": rand_hex(rng),
        "c3_para": rand_hex(rng),  # will be rendered first (in position 2)
        "c3r1_para": rand_hex(rng), "c3r1_dur": rand_hex(rng),
        "c3r2_para": rand_hex(rng), "c3r2_dur": rand_hex(rng),
        "c2_para": rand_hex(rng),  # will be rendered second (in position 3)
        "c2r_para": rand_hex(rng), "c2r_dur": rand_hex(rng),
        "c4a_para": rand_hex(rng), "c4a_dur": rand_hex(rng),
        "c4b_para": rand_hex(rng), "c4b_dur": rand_hex(rng),
    }

    dp = [rand_hex(rng) for _ in range(6)]

    body_paras = [
        make_text_para("exp5: Cases 2 and 3 swapped (tests position dependency)", dp[0], rng),
        make_text_para("Case 1 (standalone comment):", dp[1], rng),
        make_comment_para(
            "This is Case 1 with one standalone comment.",
            "one standalone comment",
            [0], [], dp[2], rng
        ),
        make_text_para("Case 3 (parent + 2 replies) — IN POSITION 2:", dp[3], rng),
        make_comment_para(
            "This is Case 3 with a parent and two replies. (in position 2)",
            "parent and two replies",
            [1], [2, 3], dp[4], rng
        ),
        make_text_para("Case 2 (parent + 1 reply) — IN POSITION 3:", dp[5], rng),
    ]
    dp2 = [rand_hex(rng) for _ in range(2)]
    body_paras += [
        make_comment_para(
            "This is Case 2 with a parent and one reply. (in position 3)",
            "parent and one reply",
            [4], [5], dp2[0], rng
        ),
        make_text_para("Case 4 (two independent comments):", dp2[1], rng),
    ]
    dp3 = [rand_hex(rng) for _ in range(1)]
    body_paras += [
        make_comment_para(
            "This is Case 4 with two independent comments.",
            "two independent comments",
            [6, 7], [], dp3[0], rng
        ),
    ]

    # Comments
    c_elems = [
        make_comment_elem(0, ids["c1_para"], "Alice", "Case 1 standalone", rng),
        make_comment_elem(1, ids["c3_para"], "Alice", "Case 3 parent (pos 2)", rng),
        make_comment_elem(2, ids["c3r1_para"], "Bob", "Case 3 reply 1 (pos 2)", rng),
        make_comment_elem(3, ids["c3r2_para"], "Bob", "Case 3 reply 2 (pos 2)", rng),
        make_comment_elem(4, ids["c2_para"], "Alice", "Case 2 parent (pos 3)", rng),
        make_comment_elem(5, ids["c2r_para"], "Bob", "Case 2 reply (pos 3)", rng),
        make_comment_elem(6, ids["c4a_para"], "Alice", "Case 4 comment A", rng),
        make_comment_elem(7, ids["c4b_para"], "Bob", "Case 4 comment B", rng),
    ]

    extended_entries = [
        (ids["c1_para"], None, False),
        (ids["c3_para"], None, False),
        (ids["c3r1_para"], ids["c3_para"], False),
        (ids["c3r2_para"], ids["c3_para"], False),
        (ids["c2_para"], None, False),
        (ids["c2r_para"], ids["c2_para"], False),
        (ids["c4a_para"], None, False),
        (ids["c4b_para"], None, False),
    ]

    ids_entries = [
        (ids["c1_para"], ids["c1_para"]),    # durableId = paraId (same bug)
        (ids["c3_para"], ids["c3_para"]),    # durableId = paraId (same bug)
        (ids["c3r1_para"], ids["c3r1_dur"]),
        (ids["c3r2_para"], ids["c3r2_dur"]),
        (ids["c2_para"], ids["c2_para"]),    # durableId = paraId (same bug)
        (ids["c2r_para"], ids["c2r_dur"]),
        (ids["c4a_para"], ids["c4a_para"]),
        (ids["c4b_para"], ids["c4b_para"]),
    ]

    doc_bytes = make_document(body_paras, rng)
    comments_bytes = make_comments_xml(c_elems)

    return (
        make_comments_extended(extended_entries),
        make_comments_ids(ids_entries),
        make_settings(doc_id),
        doc_bytes,
        comments_bytes,
    )


def build_exp6(rng):
    """
    exp6: Same structure as exp3 but ALL comments have UNIQUE durableId != paraId.
    Tests whether fixing the durableId bug makes Case 2 work.
    """
    doc_id = rand_hex(rng)

    ids = {f"c{i}_para": rand_hex(rng) for i in range(8)}
    ids.update({f"c{i}_dur": rand_hex(rng) for i in range(8)})

    cid_map = {i: (ids[f"c{i}_para"], ids[f"c{i}_dur"]) for i in range(8)}

    # Same structure as exp3:
    # c0 = Case 1 standalone
    # c1 = Case 2 parent, c2 = Case 2 reply
    # c3 = Case 3 parent, c4,c5 = Case 3 replies
    # c6,c7 = Case 4 independents
    dp = [rand_hex(rng) for _ in range(7)]

    body_paras = [
        make_text_para("exp6: All comments have unique durableId (fix durableId bug)", dp[0], rng),
        make_text_para("Case 1 (standalone):", dp[1], rng),
        make_comment_para(
            "This is Case 1 with one comment.",
            "one comment",
            [0], [], dp[2], rng
        ),
        make_text_para("Case 2 (parent + 1 reply):", dp[3], rng),
        make_comment_para(
            "This is Case 2 with one comment and one reply.",
            "one comment and one reply",
            [1], [2], dp[4], rng
        ),
        make_text_para("Case 3 (parent + 2 replies):", dp[5], rng),
        make_comment_para(
            "This is Case 3 with one comment and two replies.",
            "one comment and two replies",
            [3], [4, 5], dp[6], rng
        ),
    ]
    dp2 = [rand_hex(rng) for _ in range(2)]
    body_paras += [
        make_text_para("Case 4 (two independent):", dp2[0], rng),
        make_comment_para(
            "This is Case 4 with two independent comments.",
            "two independent comments",
            [6, 7], [], dp2[1], rng
        ),
    ]

    c_elems = [
        make_comment_elem(0, cid_map[0][0], "Alice", "Case 1 standalone", rng),
        make_comment_elem(1, cid_map[1][0], "Alice", "Case 2 parent", rng),
        make_comment_elem(2, cid_map[2][0], "Bob",   "Case 2 reply", rng),
        make_comment_elem(3, cid_map[3][0], "Alice", "Case 3 parent", rng),
        make_comment_elem(4, cid_map[4][0], "Bob",   "Case 3 reply 1", rng),
        make_comment_elem(5, cid_map[5][0], "Bob",   "Case 3 reply 2", rng),
        make_comment_elem(6, cid_map[6][0], "Alice", "Case 4 comment A", rng),
        make_comment_elem(7, cid_map[7][0], "Bob",   "Case 4 comment B", rng),
    ]

    extended_entries = [
        (cid_map[0][0], None, False),
        (cid_map[1][0], None, False),
        (cid_map[2][0], cid_map[1][0], False),  # reply → parent
        (cid_map[3][0], None, False),
        (cid_map[4][0], cid_map[3][0], False),
        (cid_map[5][0], cid_map[3][0], False),
        (cid_map[6][0], None, False),
        (cid_map[7][0], None, False),
    ]

    # KEY DIFFERENCE: ALL durableIds are unique (no durableId = paraId)
    ids_entries = [(cid_map[i][0], cid_map[i][1]) for i in range(8)]

    doc_bytes = make_document(body_paras, rng)
    comments_bytes = make_comments_xml(c_elems)

    return (
        make_comments_extended(extended_entries),
        make_comments_ids(ids_entries),
        make_settings(doc_id),
        doc_bytes,
        comments_bytes,
    )


def build_exp7(rng):
    """
    exp7: Like exp3 but with NO commentsIds.xml (just unique docId).
    Tests whether unique docId alone is sufficient (like v4 was, just with new IDs).
    Multiple paragraphs with different thread sizes.
    """
    doc_id = rand_hex(rng)

    cparas = [rand_hex(rng) for _ in range(8)]
    dp = [rand_hex(rng) for _ in range(7)]

    body_paras = [
        make_text_para("exp7: Unique docId, NO commentsIds (like v4)", dp[0], rng),
        make_text_para("Case 1 (standalone):", dp[1], rng),
        make_comment_para("Case 1 with one standalone comment.", "one standalone comment", [0], [], dp[2], rng),
        make_text_para("Case 2 (parent + 1 reply):", dp[3], rng),
        make_comment_para("Case 2 with one comment and one reply.", "one comment and one reply", [1], [2], dp[4], rng),
        make_text_para("Case 3 (parent + 2 replies):", dp[5], rng),
        make_comment_para("Case 3 with one comment and two replies.", "one comment and two replies", [3], [4, 5], dp[6], rng),
    ]
    dp2 = [rand_hex(rng) for _ in range(2)]
    body_paras += [
        make_text_para("Case 4 (two independent):", dp2[0], rng),
        make_comment_para("Case 4 with two independent comments.", "two independent comments", [6, 7], [], dp2[1], rng),
    ]

    c_elems = [
        make_comment_elem(0, cparas[0], "Alice", "Case 1 standalone", rng),
        make_comment_elem(1, cparas[1], "Alice", "Case 2 parent", rng),
        make_comment_elem(2, cparas[2], "Bob",   "Case 2 reply", rng),
        make_comment_elem(3, cparas[3], "Alice", "Case 3 parent", rng),
        make_comment_elem(4, cparas[4], "Bob",   "Case 3 reply 1", rng),
        make_comment_elem(5, cparas[5], "Bob",   "Case 3 reply 2", rng),
        make_comment_elem(6, cparas[6], "Alice", "Case 4 comment A", rng),
        make_comment_elem(7, cparas[7], "Bob",   "Case 4 comment B", rng),
    ]

    extended_entries = [
        (cparas[0], None, False),
        (cparas[1], None, False),
        (cparas[2], cparas[1], False),
        (cparas[3], None, False),
        (cparas[4], cparas[3], False),
        (cparas[5], cparas[3], False),
        (cparas[6], None, False),
        (cparas[7], None, False),
    ]

    doc_bytes = make_document(body_paras, rng)
    comments_bytes = make_comments_xml(c_elems)

    # No commentsIds
    return (
        make_comments_extended(extended_entries),
        None,  # no commentsIds
        make_settings(doc_id),
        doc_bytes,
        comments_bytes,
    )


def main():
    SCRATCH.mkdir(exist_ok=True)

    print("Generating exp4–exp7...")

    rng4 = random.Random(42)
    ex, ci, se, doc, com = build_exp4(rng4)
    write_docx(SCRATCH / "exp4-case2-isolation.docx", doc, com, ex, ci, se, include_ids=True)

    rng5 = random.Random(123)
    ex, ci, se, doc, com = build_exp5(rng5)
    write_docx(SCRATCH / "exp5-cases-swapped.docx", doc, com, ex, ci, se, include_ids=True)

    rng6 = random.Random(999)
    ex, ci, se, doc, com = build_exp6(rng6)
    write_docx(SCRATCH / "exp6-unique-durable-ids.docx", doc, com, ex, ci, se, include_ids=True)

    rng7 = random.Random(777)
    ex, ci, se, doc, com = build_exp7(rng7)
    write_docx(SCRATCH / "exp7-unique-docid-no-commentsids.docx", doc, com, ex, ci, se, include_ids=False)

    print("\nExperiments generated:")
    print("  exp4-case2-isolation.docx")
    print("    → ONLY Case 2 (parent + 1 reply) + unique docId + commentsIds")
    print("    → If fails: Case 2 has a structural bug, not context-dependent")
    print("    → If works: Case 2 was being interfered with by surrounding cases in exp3")
    print()
    print("  exp5-cases-swapped.docx")
    print("    → Cases 2 and 3 positions swapped (Case 3 in position 2, Case 2 in position 3)")
    print("    → If Case 3 now fails (in position 2): failure is POSITION-dependent")
    print("    → If Case 2 still fails (in position 3): failure is STRUCTURE-dependent (1 reply)")
    print()
    print("  exp6-unique-durable-ids.docx")
    print("    → Same as exp3 but ALL durableIds are unique (no durableId = paraId)")
    print("    → If Case 2 now works: durableId = paraId was the bug")
    print("    → If Case 2 still fails: durableId is not the issue")
    print()
    print("  exp7-unique-docid-no-commentsids.docx")
    print("    → Like v4: unique docId, NO commentsIds.xml")
    print("    → If works: unique docId alone is sufficient (commentsIds not needed)")
    print("    → If fails: v4's success was due to specific paraId values or timing")


if __name__ == "__main__":
    main()
