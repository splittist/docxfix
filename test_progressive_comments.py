"""
Progressive comment testing script.
Tests minimal comment scenarios to isolate Word compatibility issues.
"""

from pathlib import Path

from docxfix.generator import DocumentGenerator
from docxfix.validator import validate_docx
from docxfix.spec import Comment, CommentReply, DocumentSpec

out_dir = Path("scratch_out")
out_dir.mkdir(exist_ok=True)

fixtures = []

print("Generating progressive comment test files...")
print("-" * 60)

# Test 1: Minimal - single comment without reply
print("\n1. Single comment (no reply)")
single_comment = DocumentSpec(title="Single Comment", author="Test Author")
single_comment.add_paragraph(
    "This is a test sentence with an ANCHOR word.",
    comments=[
        Comment(
            text="This is a single comment.",
            anchor_text="ANCHOR",
            author="Reviewer One",
            resolved=False,
        )
    ],
)
fixtures.append(("test1-single-comment.docx", single_comment))

# Test 2: Two separate comments (no replies)
print("2. Two separate comments (no replies)")
two_comments = DocumentSpec(title="Two Comments", author="Test Author")
two_comments.add_paragraph(
    "First ANCHOR and second MARKER.",
    comments=[
        Comment(
            text="First comment.",
            anchor_text="ANCHOR",
            author="Reviewer One",
            resolved=False,
        ),
        Comment(
            text="Second comment.",
            anchor_text="MARKER",
            author="Reviewer Two",
            resolved=False,
        ),
    ],
)
fixtures.append(("test2-two-comments.docx", two_comments))

# Test 3: Single comment with one reply
print("3. Single comment with one reply")
comment_with_reply = DocumentSpec(title="Comment With Reply", author="Test Author")
comment_with_reply.add_paragraph(
    "This has an ANCHOR for threading.",
    comments=[
        Comment(
            text="Main comment.",
            anchor_text="ANCHOR",
            author="Reviewer One",
            resolved=False,
            replies=[
                CommentReply(text="Reply to main comment.", author="Reviewer Two")
            ],
        )
    ],
)
fixtures.append(("test3-comment-with-reply.docx", comment_with_reply))

# Test 4: Resolved comment
print("4. Resolved comment")
resolved_comment = DocumentSpec(title="Resolved Comment", author="Test Author")
resolved_comment.add_paragraph(
    "This has an ANCHOR for resolved comment.",
    comments=[
        Comment(
            text="This comment is resolved.",
            anchor_text="ANCHOR",
            author="Reviewer One",
            resolved=True,
        )
    ],
)
fixtures.append(("test4-resolved-comment.docx", resolved_comment))

# Test 5: Comment with multiple replies
print("5. Comment with multiple replies")
multiple_replies = DocumentSpec(title="Multiple Replies", author="Test Author")
multiple_replies.add_paragraph(
    "This has an ANCHOR for multiple replies.",
    comments=[
        Comment(
            text="Main comment with multiple replies.",
            anchor_text="ANCHOR",
            author="Reviewer One",
            resolved=False,
            replies=[
                CommentReply(text="First reply.", author="Reviewer Two"),
                CommentReply(text="Second reply.", author="Reviewer Three"),
            ],
        )
    ],
)
fixtures.append(("test5-multiple-replies.docx", multiple_replies))

# Generate all fixtures
print("\n" + "=" * 60)
print("Generating files...")
print("=" * 60)

for name, spec in fixtures:
    path = out_dir / name
    print(f"\nGenerating: {path}")
    try:
        DocumentGenerator(spec).generate(path)
        validate_docx(path)
        print(f"  ✓ Generated and validated: {name}")
    except Exception as e:
        print(f"  ✗ Error: {e}")

print("\n" + "=" * 60)
print("Done! Files saved to:", out_dir.absolute())
print("=" * 60)
print("\nNext steps:")
print("1. Open each file in Microsoft Word")
print("2. Note if Word prompts for repair")
print("3. If repair is needed, save as *-fixed.docx")
print("4. Compare generated vs. fixed files to identify issues")
