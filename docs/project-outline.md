# Project Outline (Phase 1)

This document captures the immediate direction for `docxfix` as an open-source Python library with a thin CLI for generating `.docx` fixtures, especially for legal-tech testing.

## Product direction

See also: `docs/phase-1-prd.md` for detailed Phase 1 requirements and manual-work breakdown.

- Primary form factor: **Python library first**, with a **thin CLI** wrapper.
- Purpose: generate synthetic Word documents for testing pipelines that cannot use confidential real contracts.
- Scope target: cross-runtime usage (fixtures consumed by Python, JavaScript/TypeScript, and other app stacks).

## Phase priorities

### Phase 1 (must-have)

1. **Tracked changes**
   - Insertions and deletions are top priority.
   - Should support configurable authors and basic revision metadata.
2. **Modern comments experience**
   - Threaded comments, replies, and resolved state.
   - Focus on modern comments only (legacy comments can be deferred).
3. **Complex numbering**
   - Multi-level numbering patterns suitable for legal-style clauses.
   - Two patterns: legal-list (explicit `numPr` per paragraph) and heading-based/styled (numbering linked via style definitions).

### Phase 2 (deferred)

- Sections (multiple sections, orientation and header/footer variation).

## Data generation approach

- Use synthetic placeholder text for now (e.g., lorem ipsum).
- Faker-driven metadata (names, initials, etc.) is acceptable and desirable.
- Deterministic generation should be supported where practical (e.g., random seed), but **byte-identical output is not required** due to Word/environmental normalization behavior.
- Reference OOXML schemas in `./schemas` for validation and structural guidance.
- Maintain a corpus of `.docx` fixtures with sidecar `.md` descriptions in `./corpus`.

## Architecture outline

- Keep a strongly typed internal spec model representing fixture intent.
- Treat `.docx` files as Open Packaging Convention archives with coordinated XML parts.
- Implement feature-specific mutation modules (tracked changes, comments, numbering).
- Add validation that combines XML/schema checks and semantic integrity checks.

## Testing strategy

### Cross-platform core CI

- Unit tests for XML mutators and helpers.
- Integration tests generating fixture files and asserting expected feature counts/structures.
- Snapshot-style tests where stable and useful.

### Windows compatibility lane

- Optional lane that opens generated fixtures in Microsoft Word and verifies they remain valid after save.
- This lane is a compatibility check only; core CI remains cross-platform.

## BDD integration intent

- The library/CLI should be able to consume test-friendly fixture specs that map well from BDD scenario tables.
- Output artifacts should be easy to use from non-Python stacks (e.g., TypeScript test suites).

## Non-goals (for now)

- Compliance/audit guarantees.
- Confidentiality certifications or governance workflows.
