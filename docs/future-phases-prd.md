# Future Scope PRD (Lean)

## 1. Purpose

This document defines the next scope after Phase 2. The goal is intentionally narrow: generate simple `.docx` fixtures from BDD-style descriptions for legal-tech testing.

## 2. Scope guardrails

- Focus on fixture generation only.
- Keep workflows simple for a single user/small team.
- Defer analysis mode, assertion engines, and formal versioning policy until demand exists.

## 3. Phase 3 PRD: BDD Fixture Generation

## 3.1 Problem statement

Current usage is strongest from Python code. For BDD workflows, users need a straightforward way to describe fixtures in text/tables and generate `.docx` files via CLI.

## 3.2 Goals

1. Support fixture input files in JSON/YAML.
2. Add CLI generation from spec files and simple batch manifests.
3. Provide a lightweight mapping helper from BDD row fields to fixture spec fields.
4. Ship clear examples for Python and JS/TS test suites consuming generated files.

## 3.3 Non-goals

- Analyze existing `.docx` files.
- Built-in assertion/reporting framework.
- Formal compatibility/versioning contract.
- Full Gherkin parser/runtime.

## 3.4 Functional requirements

### FR-3.1 External fixture input format

- Accept JSON and YAML fixture descriptions.
- Keep fields aligned with existing `DocumentSpec` concepts.
- Validate required fields and return actionable error messages.

Acceptance criteria:
- At least 8 example fixture specs are included.
- Invalid input errors include field path and reason.

### FR-3.2 CLI spec and batch generation

- Support:
  - `docxfix create --spec path/to/spec.yml --out out.docx`
  - `docxfix batch --manifest path/to/fixtures.yml --out-dir fixtures/`
- Manifest supports fixture id and output filename.
- Support seeded deterministic output from input spec.

Acceptance criteria:
- Batch run handles at least 20 fixtures.
- CLI exits non-zero when one or more fixtures fail.
- Failure output identifies fixture id and input file.

### FR-3.3 BDD row mapping helper

- Provide a small utility that maps table row keys to fixture fields.
- Support a constrained alias set for common cases:
  - tracked changes on/off,
  - comment thread count,
  - numbering depth,
  - section usage on/off.
- Keep mapping explicit and documented.

Acceptance criteria:
- Mapping helper has unit tests for happy paths and invalid aliases.
- Example shows row -> spec -> `.docx` in one flow.

### FR-3.4 Documentation and examples

- Update guide with non-Python-first workflow.
- Add minimal BDD examples:
  - pytest-bdd-style table usage,
  - JS/TS test flow that shells out to CLI.

Acceptance criteria:
- Examples are runnable and covered in CI smoke checks.
- Guide includes a “quick start” path under 10 minutes.

## 3.5 Deliverables

1. External JSON/YAML input support.
2. CLI `create --spec` and `batch` workflow.
3. BDD row mapping helper.
4. Updated docs and runnable examples.

## 3.6 Milestones

1. M3.1: Input parser + validation errors. ✅ **Complete** — `input_parser.py`, 64 tests, 8 example specs.
2. M3.2: CLI `--spec` and batch manifest flow.
3. M3.3: BDD row mapping helper + tests.
4. M3.4: Docs/examples cleanup.

## 4. Deferred backlog (only if needed later)

- Analysis mode for existing `.docx` files.
- Built-in assertion profiles/reporting.
- Formal versioning and compatibility policy.

## 5. Success criteria

1. BDD description -> `.docx` generation works from CLI without Python coding.
2. Errors are actionable for malformed fixture inputs.
3. A small set of realistic legal-tech fixture scenarios can be generated repeatably in CI.
