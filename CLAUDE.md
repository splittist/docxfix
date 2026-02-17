# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

docxfix is a Python library (with a thin CLI) for creating `.docx` fixture files with specific characteristics (tracked changes, modern threaded comments, multilevel numbering, sections) for testing document processing systems — primarily legal-tech contract-analysis workflows. It generates valid OOXML ZIP archives using lxml.

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

1. **`spec.py`** — Pure-Python dataclass tree (`DocumentSpec`, `Paragraph`, `TrackedChange`, `Comment`, `CommentReply`, `NumberedParagraph`). `DocumentSpec.add_paragraph()` returns `self` for fluent chaining. Paragraphs support `heading_level` for styled numbering.
2. **`constants.py`** — `NAMESPACES`, `WORD_NAMESPACES` (30+ entries), and XML string constants extracted from the generator.
3. **`boilerplate.py`** — Stateless OOXML part generators (content types, relationships, core properties).
4. **`generator.py`** — `DocumentGenerator` converts a spec into OOXML parts as lxml trees and writes them into a ZIP. Delegates feature-specific logic to `parts/` modules. Supports deterministic output via seeded `random.Random` instance.
5. **`parts/`** — Feature modules extracted from the generator monolith:
   - `context.py` — `GeneratorContext` dataclass (shared state, seeded RNG, ID generation).
   - `comments.py` — Comment and reply XML generation (comments.xml, commentsExtended.xml).
   - `tracked_changes.py` — Insertion/deletion markup with positional interleaving.
   - `numbering.py` — Legal-list and heading-based multilevel numbering (numbering.xml).
   - `styles.py` — ListParagraph and Heading1–4 style definitions (styles.xml).
   - `sections.py` — Section breaks, headers, and footers.
6. **`validator.py`** — Post-generation ZIP structure, XML well-formedness, and semantic checks (ID uniqueness, anchor integrity, relationship completeness, content type coverage).
7. **`xml_utils.py`** — Small lxml helpers.
8. **`cli.py`** — Typer CLI wrapping the above.

XML elements use Clark notation (`{namespace}localname`). `WORD_NAMESPACES` is used as `nsmap=` on root elements to satisfy Word's namespace expectations.

## Project Status

- **Phase 1 (complete):** Tracked changes, modern threaded comments, legal-list + heading-based numbering, core spec/generator/validator, corpus, docs.
- **Phase 2 (complete):** Refactored generator into `constants.py`, `boilerplate.py`, and `parts/` modules. Added deterministic seed support, expanded semantic validation, snapshot tests (syrupy), sections/headers/footers. 116 tests passing.
- **Phase 3 (next):** BDD fixture generation — external JSON/YAML input format, CLI `--spec`/`batch` workflows, BDD row mapping helper, runnable examples for Python and JS/TS consumers. See `docs/future-phases-prd.md`.
- **Deferred backlog:** Analysis mode for existing `.docx` files, assertion/reporting framework, formal versioning policy.

## Key Conventions

- Python 3.12+ features (PEP 695 `type` statements, `str | None` unions)
- Ruff for linting/formatting (line-length 88, double quotes)
- Tests use `tempfile.TemporaryDirectory()`, parse generated ZIPs with lxml, and assert via XPath with explicit namespace maps
- Snapshot tests use syrupy for stable normalized XML output
- Deterministic mode: pass `seed=` to `DocumentGenerator` for byte-identical output in CI
- `corpus/` contains curated golden `.docx` fixtures with sidecar `.md` descriptions

## Workflow

Source-of-truth docs: `docs/project-outline.md`, `docs/future-phases-prd.md`, `PROGRESS.txt`. When implementing features: pick highest-priority item from the current phase PRD, implement with tests, verify tests pass, update PRD, append to PROGRESS.txt, commit. Work on one item at a time.
