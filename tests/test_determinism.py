"""Tests for deterministic document generation."""

import tempfile
from pathlib import Path

from docxfix.generator import DocumentGenerator
from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    NumberedParagraph,
    TrackedChange,
)


def _generate_bytes(spec: DocumentSpec) -> bytes:
    """Generate a docx and return its raw bytes."""
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(path)
        return path.read_bytes()


def _make_seeded_spec(seed: int = 42) -> DocumentSpec:
    """Build a spec exercising all features with a fixed seed."""
    spec = DocumentSpec(seed=seed)
    spec.add_paragraph("Simple text")
    spec.add_paragraph(
        "Numbered item",
        numbering=NumberedParagraph(level=0),
    )
    spec.add_paragraph("Chapter One", heading_level=1)
    spec.add_paragraph(
        "Text with comment",
        comments=[
            Comment(
                text="A comment",
                anchor_text="comment",
                replies=[CommentReply(text="A reply")],
            )
        ],
    )
    spec.add_paragraph(
        "Hello cruel world",
        tracked_changes=[
            TrackedChange(
                change_type=ChangeType.DELETION,
                text="cruel ",
            ),
            TrackedChange(
                change_type=ChangeType.INSERTION,
                text="beautiful ",
                insert_after="Hello ",
            ),
        ],
    )
    return spec


def test_seeded_generation_is_byte_identical():
    """Two generations with the same seed produce identical output."""
    spec1 = _make_seeded_spec(seed=42)
    spec2 = _make_seeded_spec(seed=42)
    assert _generate_bytes(spec1) == _generate_bytes(spec2)


def test_different_seeds_produce_different_output():
    """Different seeds produce different output."""
    spec1 = _make_seeded_spec(seed=42)
    spec2 = _make_seeded_spec(seed=99)
    assert _generate_bytes(spec1) != _generate_bytes(spec2)


def test_unseeded_does_not_pollute_global_random():
    """Unseeded generation uses an isolated RNG, not module-level random."""
    import random

    random.seed(12345)
    before = random.random()
    random.seed(12345)

    spec = DocumentSpec()
    spec.add_paragraph("Test")
    _generate_bytes(spec)

    after = random.random()
    assert before == after
