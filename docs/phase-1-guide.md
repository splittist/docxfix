# Phase 1 User and Developer Guide

## Overview

`docxfix` is a Python library and CLI for generating synthetic `.docx` fixture files with specific characteristics for testing document processing systems. It produces valid OOXML ZIP archives targeting modern Microsoft Word.

### Supported features

- **Tracked changes** — insertions and deletions with author/date metadata.
- **Modern threaded comments** — top-level comments, replies, and resolved state.
- **Complex numbering** — legal-style multilevel lists and heading-based (styled) numbering.
- **Combined scenarios** — all features can be used together in the same document.

## Installation

Requires Python 3.12+.

```bash
# Install with dev dependencies (editable)
uv pip install -e ".[dev]"
```

## CLI Usage

### Generate a fixture

```bash
uv run docxfix create output.docx
```

Options:

| Flag | Description |
|------|-------------|
| `--validate / --no-validate` | Run post-generation validation (default: on) |
| `--verbose, -v` | Enable verbose output |
| `--template, -t` | Template `.docx` to use as base (reserved) |

### Show version info

```bash
uv run docxfix info
```

## Python API

The primary workflow is **Spec → Generator → ZIP**.

### 1. Build a spec

```python
from docxfix.spec import (
    DocumentSpec,
    TrackedChange,
    ChangeType,
    Comment,
    CommentReply,
    NumberedParagraph,
)

spec = DocumentSpec(title="Contract Draft", author="Jane Doe")
```

### 2. Add paragraphs

`DocumentSpec.add_paragraph()` returns `self` for fluent chaining.

#### Plain text

```python
spec.add_paragraph("This is a normal paragraph.")
```

#### Tracked changes

```python
from datetime import datetime

# Deletion: "the important" is marked as deleted within the paragraph text
spec.add_paragraph(
    "Review the important contract terms.",
    tracked_changes=[
        TrackedChange(
            change_type=ChangeType.DELETION,
            text="important ",
            author="Editor",
            date=datetime(2025, 3, 15, 10, 0),
            revision_id=2,
        ),
    ],
)

# Insertion: "additional " is inserted after "Review the"
spec.add_paragraph(
    "Review the contract terms.",
    tracked_changes=[
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="additional ",
            author="Editor",
            date=datetime(2025, 3, 15, 10, 30),
            revision_id=3,
            insert_after="Review the ",
        ),
    ],
)
```

**TrackedChange fields:**

| Field | Type | Description |
|-------|------|-------------|
| `change_type` | `ChangeType` | `INSERTION` or `DELETION` |
| `text` | `str` | The inserted or deleted text |
| `author` | `str` | Revision author (default: `"Test User"`) |
| `date` | `datetime \| None` | Revision timestamp (default: now) |
| `revision_id` | `int` | Revision ID value (default: `1`) |
| `insert_after` | `str` | For insertions: substring after which to insert (default: `""` = append) |

**Positioning rules:**

- **Deletions:** `text` must be a substring of the paragraph text. The generator locates it and wraps it in `<w:del>`.
- **Insertions:** `insert_after` identifies the anchor point. The insertion is placed immediately after that substring. If empty, the insertion is appended at the end.

#### Comments

```python
spec.add_paragraph(
    "The parties agree to the following terms.",
    comments=[
        Comment(
            text="Need legal review on this clause.",
            anchor_text="following terms",
            author="Reviewer",
            replies=[
                CommentReply(
                    text="Reviewed and approved.",
                    author="Legal",
                ),
            ],
            resolved=False,
        ),
    ],
)
```

**Comment fields:**

| Field | Type | Description |
|-------|------|-------------|
| `text` | `str` | Comment body text |
| `anchor_text` | `str` | Substring in the paragraph to anchor the comment to |
| `author` | `str` | Comment author (default: `"Test User"`) |
| `date` | `datetime \| None` | Comment timestamp (default: now) |
| `replies` | `list[CommentReply]` | Reply chain |
| `resolved` | `bool` | Whether the comment thread is resolved (default: `False`) |

#### Numbering — legal-style lists

```python
spec.add_paragraph("Definitions", numbering=NumberedParagraph(level=0))
spec.add_paragraph("General terms", numbering=NumberedParagraph(level=1))
spec.add_paragraph("Specific terms", numbering=NumberedParagraph(level=2))
```

Produces legal-style numbering: `1.`, `1.1.`, `1.1.1.`, etc. Supports levels 0–8.

**NumberedParagraph fields:**

| Field | Type | Description |
|-------|------|-------------|
| `level` | `int` | Nesting level, 0-based (default: `0`) |
| `numbering_id` | `int` | Numbering definition ID (default: `1`) |

#### Numbering — heading-based (styled)

```python
spec.add_paragraph("Article I — Definitions", heading_level=1)
spec.add_paragraph("Section 1.1 — Scope", heading_level=2)
spec.add_paragraph("Subsection (a)", heading_level=3)
```

Heading levels 1–4 map to `Heading1`–`Heading4` styles. Numbering is linked through style definitions rather than explicit `numPr` on each paragraph.

#### Combined features

All features can be used together on the same paragraph:

```python
spec.add_paragraph(
    "The seller shall deliver the goods.",
    tracked_changes=[
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="promptly ",
            author="Editor",
            insert_after="shall ",
        ),
    ],
    comments=[
        Comment(
            text="Timeframe needs clarification.",
            anchor_text="deliver the goods",
            author="Reviewer",
        ),
    ],
    numbering=NumberedParagraph(level=0),
)
```

### 3. Generate the document

```python
from docxfix.generator import DocumentGenerator

generator = DocumentGenerator(spec)
generator.generate("output.docx")
```

### 4. Validate (optional)

```python
from docxfix.validator import validate_docx, ValidationError

try:
    validate_docx("output.docx")
except ValidationError as e:
    print(f"Validation failed: {e}")
```

Validation checks:

- ZIP archive structure (required OOXML entries present).
- XML well-formedness of all XML parts.

## Spec Examples

### Tracked-change-heavy contract

```python
spec = DocumentSpec(title="Contract v2", author="Legal Team")
spec.add_paragraph(
    "The Client agrees to pay the agreed amount upon delivery.",
    tracked_changes=[
        TrackedChange(
            change_type=ChangeType.DELETION,
            text="agreed amount",
            author="Counsel A",
            revision_id=1,
        ),
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="total fee of $50,000",
            author="Counsel A",
            revision_id=2,
            insert_after="pay the ",
        ),
    ],
)
spec.add_paragraph(
    "Services shall commence on the effective date.",
    tracked_changes=[
        TrackedChange(
            change_type=ChangeType.INSERTION,
            text="professional ",
            author="Counsel B",
            revision_id=3,
            insert_after="Services shall commence ",
        ),
    ],
)
```

### Comment-threaded review document

```python
spec = DocumentSpec(title="Draft Agreement", author="Author")
spec.add_paragraph(
    "This Agreement is entered into as of the Effective Date.",
    comments=[
        Comment(
            text="Should we specify the exact date?",
            anchor_text="Effective Date",
            author="Reviewer A",
            replies=[
                CommentReply(text="Yes, let's pin it to signing.", author="Author"),
                CommentReply(text="Agreed.", author="Reviewer A"),
            ],
            resolved=True,
        ),
    ],
)
```

### Legal clause hierarchy

```python
spec = DocumentSpec(title="Terms of Service")
spec.add_paragraph("Definitions", numbering=NumberedParagraph(level=0))
spec.add_paragraph('"Service" means the hosted platform.', numbering=NumberedParagraph(level=1))
spec.add_paragraph('"User" means any registered individual.', numbering=NumberedParagraph(level=1))
spec.add_paragraph("Obligations", numbering=NumberedParagraph(level=0))
spec.add_paragraph("The User shall not:", numbering=NumberedParagraph(level=1))
spec.add_paragraph("reverse-engineer the Service;", numbering=NumberedParagraph(level=2))
spec.add_paragraph("share access credentials.", numbering=NumberedParagraph(level=2))
```

## Architecture

```
DocumentSpec  →  DocumentGenerator  →  .docx (ZIP)
  (spec.py)       (generator.py)
                       ↓
                  validator.py (post-generation checks)
```

**Key modules:**

| Module | Responsibility |
|--------|---------------|
| `spec.py` | Pure-Python dataclass tree defining fixture intent |
| `generator.py` | Converts spec into OOXML parts (lxml trees) and writes ZIP |
| `validator.py` | Post-generation ZIP structure and XML checks |
| `xml_utils.py` | Small lxml helpers |
| `cli.py` | Typer CLI wrapping the above |

**Generated OOXML parts** (conditional on spec content):

| Part | When included |
|------|--------------|
| `word/document.xml` | Always |
| `word/comments.xml` | When paragraphs have comments |
| `word/commentsExtended.xml` | When paragraphs have comments |
| `word/numbering.xml` | When paragraphs have numbering or heading levels |
| `word/styles.xml` | When paragraphs have numbering or heading levels |
| `word/settings.xml` | Always |
| `word/webSettings.xml` | Always |
| `word/fontTable.xml` | Always |
| `word/theme/theme1.xml` | Always |
| `docProps/app.xml` | Always |

## Golden Corpus

The `corpus/` directory contains 7 curated `.docx` fixtures authored in Word, each with a sidecar `.md` description:

| Fixture | Feature |
|---------|---------|
| `single-insertion.docx` | One tracked insertion |
| `single-deletion.docx` | One tracked deletion |
| `mixed-insert-delete.docx` | Insertion + deletion in one paragraph |
| `comment-thread.docx` | Comment with reply |
| `resolved-comment.docx` | Resolved comment thread |
| `legal-list.docx` | Multilevel legal numbering |
| `styled-numbering.docx` | Heading-based numbering via styles |

See `corpus/README.md` for the sidecar format specification.

## Known Limitations

### Tracked changes

- Deletion text must be an exact substring of the paragraph text.
- Insertion positioning requires specifying `insert_after` as an exact substring.
- Overlapping tracked changes within the same paragraph are not validated for conflicts.

### Comments

- One comment per anchor text range per paragraph. Multiple comments on the same paragraph use separate anchor ranges.
- Anchor text must be an exact substring match (no fuzzy/regex matching).
- Reply threading is single-level (replies link to the top-level comment, not to other replies).

### Numbering

- Legal-list numbering uses a fixed format pattern (decimal, `%1.%2.` style). Custom format strings are not supported.
- Heading-based numbering supports levels 1–4 only.
- Numbering restart across sections is not supported (single continuous sequence).

### General

- No section layout support (orientation, headers/footers, page breaks).
- No legacy comments mode (modern threaded comments only).
- Output is not byte-identical across runs (timestamps and IDs vary).
- The `seed` field on `DocumentSpec` is reserved but not yet wired to deterministic generation.
- CLI generates a hardcoded demo document; custom specs require the Python API.

## Compatibility Checklist

### Automated (CI)

- [x] ZIP structure validation passes for all generated fixtures.
- [x] XML well-formedness checks pass for all parts.
- [x] Generated files contain expected `<w:ins>`, `<w:del>` counts.
- [x] Comment thread graph (comments.xml + commentsExtended.xml) is internally consistent.
- [x] Numbering definitions and paragraph references are coherent.
- [x] Integration tests cover all feature families and combined scenarios.
- [x] 92 tests passing across unit and integration suites.

### Manual (Word verification)

- [ ] Generated tracked-change documents open in Word without repair prompts.
- [ ] Insertions and deletions display correctly in Review mode.
- [ ] Comment threads display with proper nesting in Word's comment pane.
- [ ] Resolved comments show resolved state.
- [ ] Legal-list numbering renders as expected (1., 1.1., 1.1.1.).
- [ ] Heading-based numbering renders with correct heading styles.
- [ ] Combined-feature documents (tracked changes + comments + numbering) open cleanly.

### Target environment

- Microsoft Word (current/up-to-date versions only).
- Word for Windows and Word for Mac are both in scope.
- Word Online is out of scope for Phase 1.
