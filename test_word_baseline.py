"""
Create a minimal baseline file with one comment for manual testing.
User will:
1. Open this in Word
2. Save a version with a reply added -> test-baseline-with-reply.docx
3. Save a version with comment marked resolved -> test-baseline-resolved.docx
"""

from pathlib import Path
from docxfix.spec import DocumentSpec, Comment
from docxfix.generator import DocumentGenerator
from docxfix.validator import validate_docx

# Create minimal document with one comment
spec = DocumentSpec(title="Baseline Test", author="Test Author")
spec.add_paragraph(
    "This is a test document with a single COMMENT.",
    comments=[
        Comment(
            text="This is the main comment.",
            anchor_text="COMMENT",
            author="Reviewer One",
            resolved=False,
        )
    ],
)

# Generate the document
output_path = Path("scratch_out/test-baseline.docx")
DocumentGenerator(spec).generate(output_path)

# Validate
validate_docx(output_path)

print(f"✓ Generated: {output_path}")
print()
print("Next steps:")
print("1. Open test-baseline.docx in Word")
print("2. Add a reply to the comment and save as: scratch_out/test-baseline-with-reply.docx")
print("3. Open test-baseline.docx again, mark the comment as resolved")
print("   (right-click comment -> Resolve) and save as: scratch_out/test-baseline-resolved.docx")
print()
print("Then we can compare the XML to see what Word generates differently.")
