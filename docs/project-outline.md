# Project Outline

This document captures the practical direction for `docxfix` as an open-source Python library with a thin CLI for generating `.docx` fixtures for legal-tech test automation.

## Product direction

- Primary form factor: **Python library first**, with a **thin CLI** wrapper.
- Purpose: help teams building contract-analysis systems generate representative Word fixtures from simple BDD-style descriptions.
- Scope target: cross-runtime usage (fixtures consumed by Python, JavaScript/TypeScript, and other app stacks).
- See also:
  - `docs/phase-1-prd.md`
  - `docs/phase-2-prd.md`
  - `docs/future-phases-prd.md`

## Current status

- Phase 1: completed (tracked changes, modern comments, legal numbering baseline).
- Phase 2: completed (refactor/hardening, deterministic generation, broader validation, sections support in implementation).
- Next work: Phase 3 focused on BDD interchange and CLI fixture generation only.

## Roadmap

### Phase 1 (completed)

1. Tracked changes.
2. Modern threaded comments.
3. Complex numbering (legal-list + heading/styled).
4. Core typed spec, generator, validation, and baseline docs/corpus.

### Phase 2 (completed)

1. Code quality hardening and modularization.
2. Deterministic generation with fixed seed support.
3. Expanded semantic validation and snapshot coverage.
4. Sections/header/footer implementation and documentation alignment.

### Phase 3 (next: BDD fixture generation)

1. External fixture spec format for non-Python users (JSON/YAML).
2. CLI support for loading fixture specs and batch generation.
3. First-class BDD mapping helpers (scenario-table row -> fixture spec).
4. Clear examples for using generated fixtures in BDD test suites.

### Backlog (deferred until needed)

- Formal versioning/compatibility policy.
- Analyzer/assertion workflows.
- Broader adoption and packaging work.

## Data approach

- Synthetic placeholder text remains default.
- Faker-style metadata generation remains acceptable.
- Deterministic mode (seeded) remains a core capability for CI reproducibility.
- OOXML schemas in `./schemas` remain the primary structural reference.
- `./corpus` remains the curated source of golden fixtures with sidecar markdown.

## Architecture direction

- Keep a strongly typed internal Python spec model as the canonical representation.
- Add simple interchange input support for external users (JSON/YAML).
- Keep OOXML mutation logic modular by feature family.
- Maintain a two-layer quality bar:
  - structural checks (XML/schema/package),
  - semantic checks (IDs, anchors, relationships, numbering/section integrity).

## Testing direction

### Cross-platform core CI

- Unit tests for spec parsing and generation helpers.
- Integration tests for BDD description -> `.docx` generation round trips.
- Snapshot tests for stable normalized outputs.

### Windows compatibility lane

- Optional Word-open and save round-trip checks.
- Compatibility lane remains supplementary; core CI remains cross-platform.

## Non-goals

- Enterprise document governance workflows.
- Compliance, audit, or legal defensibility guarantees.
- Advanced template-authoring UX (focus stays on test automation primitives).
