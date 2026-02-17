# Phase 2: Code Quality & Hardening — Deliverables

## What was delivered

### M1: Extract constants and boilerplate

Moved static data out of the 1618-line `generator.py` monolith into focused modules:

- **`constants.py`** — `NAMESPACES`, `WORD_NAMESPACES`, and XML string constants (`SETTINGS_XML`, `WEB_SETTINGS_XML`, `FONT_TABLE_XML`, `APP_PROPERTIES_XML`, `THEME_XML_B64`).
- **`boilerplate.py`** — Stateless functions for settings, web settings, footnotes, endnotes, font table, theme, core properties, and app properties.

### M2: Extract feature modules

Split feature-specific methods into a `parts/` subpackage. Each module receives explicit parameters via a shared `GeneratorContext` dataclass instead of reaching into `self`:

- **`parts/context.py`** — `GeneratorContext` with revision/comment counters, comment metadata accumulator, and seeded RNG instance.
- **`parts/comments.py`** — Comment anchoring, `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`.
- **`parts/tracked_changes.py`** — Tracked change interleaving (plain text, deletions, insertions) and combined comment+TC handling.
- **`parts/numbering.py`** — Legal-list and heading abstract numbering definitions.
- **`parts/styles.py`** — Heading styles and base style definitions.
- **`parts/sections.py`** — Section normalization, header/footer part manifest, section properties.

`generator.py` reduced from 1618 lines to ~620 lines.

### M3: Enhanced validation

Extended `validator.py` with 5 semantic checks:

1. Comment ID uniqueness in `comments.xml`.
2. Tracked change ID uniqueness in `document.xml`.
3. Comment anchor integrity (`commentRangeStart`/`commentRangeEnd` pairing).
4. Relationship completeness (every `rId` in `document.xml` exists in rels).
5. Content type coverage (every ZIP part has a matching content type).

Added 7 new tests in `test_validator.py`.

### M4: Snapshot tests with syrupy

5 snapshot tests covering key XML parts with fixed seed (42):

- `document.xml` for a simple paragraph.
- `comments.xml` for a comment with reply.
- `numbering.xml` for legal-list numbering.
- `styles.xml` for heading numbering.
- `[Content_Types].xml` with all features enabled.

### M5: Wire deterministic seed

Made generation fully deterministic when `DocumentSpec.seed` is set:

- `GeneratorContext` uses a `random.Random(seed)` instance instead of module-level `random`.
- Spec objects with default `datetime.now()` are overwritten with a fixed reference datetime (`2024-01-01T00:00:00Z`).
- `create_core_properties()` accepts an optional fixed timestamp.
- Two `DocumentGenerator(spec_with_seed).generate()` calls produce byte-identical output.
- Unseeded generation does not pollute global `random` state.

### M6: Cleanup and docs update

- Removed unused `Path` import from `cli.py`.
- Updated `docs/phase-1-guide.md`: added sections to supported features, section API examples, updated architecture table and validation description, removed stale limitations about missing section support and seed not being wired.
- Created `docs/phase-2-prd.md` (this document).
- Updated `PROGRESS.txt` with Phase 2 entries.

## Test results

116 tests passing (101 from Phase 1 + 7 validation + 5 snapshot + 3 determinism).

## Deferred items for Phase 3

- Spec-level validation (section bounds, deletion text existence, anchor text existence).
- CLI support for custom specs (currently generates a hardcoded demo document).
- Legacy comments mode.
- Image/embedded object support.
