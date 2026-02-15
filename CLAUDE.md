# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

docxfix is a Python CLI tool for creating `.docx` fixture files with specific characteristics (tracked changes, comments, numbering) for testing document processing systems. It generates valid OOXML ZIP archives using lxml.

## Commands

```bash
# Install (editable, with dev deps)
uv pip install -e ".[dev]"

# Run CLI
uv run docxfix create output.docx
uv run docxfix info

# Tests
uv run pytest                                          # all tests
uv run pytest tests/test_generator.py                  # single file
uv run pytest -k "test_generator_simple_comment"       # single test by name

# Lint and format
ruff check src/ tests/
ruff format src/ tests/
```

## Architecture

**Spec → Generator → ZIP** pipeline:

1. **`spec.py`** — Pure-Python dataclass tree (`DocumentSpec`, `Paragraph`, `TrackedChange`, `Comment`, `CommentReply`, `NumberedParagraph`). `DocumentSpec.add_paragraph()` returns `self` for fluent chaining.
2. **`generator.py`** — `DocumentGenerator` converts a spec into OOXML parts as lxml trees and writes them into a ZIP. Conditional parts (`has_comments`, `has_numbering`) control which XML entries are emitted.
3. **`validator.py`** — Post-generation ZIP structure and XML well-formedness checks.
4. **`xml_utils.py`** — Small lxml helpers.
5. **`cli.py`** — Typer CLI wrapping the above.

XML elements use Clark notation (`{namespace}localname`). A `NAMESPACES` dict holds core URIs; `WORD_NAMESPACES` (30+ entries) is used as `nsmap=` on root elements to satisfy Word's namespace expectations.

## Key Conventions

- Python 3.12+ features (PEP 695 `type` statements, `str | None` unions)
- Ruff for linting/formatting (line-length 88, double quotes)
- Tests use `tempfile.TemporaryDirectory()`, parse generated ZIPs with lxml, and assert via XPath with explicit namespace maps
- `corpus/` contains curated golden `.docx` fixtures with sidecar `.md` descriptions

## Workflow (from PROMPT.md)

Source-of-truth docs: `docs/project-outline.md`, `docs/phase-1-prd.md`, `PROGRESS.txt`. When implementing features: pick highest-priority item, implement with tests, verify tests pass, update PRD, append to PROGRESS.txt, commit. Work on one item at a time.
