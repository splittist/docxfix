# Phase 1 Product Requirements Document (PRD)

## 1. Purpose

Phase 1 delivers the first production-usable version of `docxfix` focused on generating realistic `.docx` fixtures for legal-tech and document-processing tests.

This PRD expands the high-level outline in `docs/project-outline.md` into concrete scope, deliverables, acceptance criteria, and manual tasks required to launch Phase 1.

## 2. Problem Statement

Teams need representative Word documents with tracked changes, comments, and complex numbering, but cannot use confidential client contracts in test suites.

Phase 1 addresses this by providing deterministic-enough fixture generation from a structured spec via Python API and thin CLI.

## 3. Goals and Non-Goals

### 3.1 Goals

1. Generate valid `.docx` fixtures including:
   - tracked insertions/deletions,
   - modern threaded comments (including reply chains and resolved state),
   - legal-style multilevel numbering.
2. Provide a typed internal spec model that cleanly maps fixture intent to XML mutations.
3. Offer a thin CLI to generate fixtures from spec input suitable for CI and BDD pipelines.
4. Establish confidence with automated validation and a curated manual "golden" corpus.

### 3.2 Non-Goals (Phase 1)

- Section layout complexity (orientation changes, section-specific headers/footers).
- Legacy comments mode parity.
- Guarantees of byte-identical outputs across Word versions/runtimes.
- Compliance/security certifications.

## 4. Target Users and Core Use Cases

### 4.1 Primary users

- QA/test engineers in legal-tech platforms.
- Backend/infrastructure developers building parser and ingestion tests.
- Cross-runtime consumers (Python and JS/TS tests) that need fixture artifacts.

### 4.2 Core use cases

1. **Generate tracked-change-heavy contracts** for parser regression tests.
2. **Generate comment-threaded documents** to test extraction of conversation context.
3. **Generate numbering-heavy legal clauses** for clause segmentation and numbering integrity tests.
4. **Produce deterministic fixture sets** from seed-based inputs for repeatable CI behavior.

## 5. Functional Requirements

## FR-1: Tracked Changes (Insertions/Deletions)

### Required capabilities

- Add insertion runs (`<w:ins>`) and deletion runs (`<w:del>`) in body content.
- Configure revision metadata at minimum:
  - author,
  - date/timestamp,
  - revision id-like value where applicable.
- Support multiple tracked changes in one paragraph and across paragraphs.
- Ensure generated structure remains openable in Word.

### Acceptance criteria

- Generated docs contain expected counts of `<w:ins>` and `<w:del>` matching spec intent.
- Each tracked change contains required metadata fields.
- Files pass XML/schema checks and Word-open smoke checks (in compatibility lane/manual pass).

## FR-2: Modern Threaded Comments

### Required capabilities

- Create modern comment threads:
  - top-level comment,
  - one or more replies,
  - resolved/unresolved state.
- Anchor comments to specific run ranges in main document content.
- Support configurable author identity metadata.

### Acceptance criteria

- Comment thread graph in generated package is internally consistent.
- Reply relationships and resolved state serialize correctly.
- Word opens generated files and displays thread structure without repair prompts.

## FR-3: Complex Numbering

### Required capabilities

- Generate multilevel numbering definitions suitable for legal structures (e.g., 1., 1.1, 1.1(a), etc.).
- Apply numbering across paragraphs with nesting depth changes.
- Support restarting numbering sequences where specified.

### Acceptance criteria

- `numbering.xml` definitions and paragraph references are coherent.
- Paragraph order reflects expected legal clause hierarchy.
- Generated documents remain valid and render numbering as intended in Word.

## FR-4: API and CLI Surface

### Required capabilities

- Python API for loading/constructing fixture specs and producing `.docx` output.
- Thin CLI command(s) to generate output artifacts from spec files.
- Seed option for predictable random metadata/text generation.
- Helpful error messages for invalid specs.

### Acceptance criteria

- CLI can generate at least one fixture per Phase 1 feature family.
- CLI/API expose enough options to support CI usage.
- Invalid spec errors identify the failing field/path.

## FR-5: Validation and Integrity

### Required capabilities

- XML well-formedness checks on touched parts.
- Schema validation where feasible (with documented exceptions if schemas are partial).
- Semantic integrity checks (e.g., references between parts, ids, relationships).
- Use the OOXML schema set in `./schemas` as the primary validation reference.

### Acceptance criteria

- Validation pipeline fails fast on structural and relational inconsistencies.
- Test coverage includes representative invalid cases for each feature family.

## 6. Deliverables

1. **Feature implementation modules** for tracked changes, comments, numbering.
2. **Typed spec model** and schema/contracts for fixture definitions.
3. **CLI commands** for fixture generation.
4. **Automated tests** (unit + integration) for core logic.
5. **Golden corpus** of `.docx` fixtures with sidecar `.md` descriptions stored in `./corpus`.
6. **Validation subsystem** (schema + semantic checks).
7. **Phase 1 docs**:
   - usage guide,
   - spec examples,
   - known limitations,
   - compatibility checklist.
   - corpus format guide in [corpus/README.md](corpus/README.md).

## 7. Milestones

### M1: Spec + validation foundation

- Finalize Phase 1 spec fields.
- Implement initial parser/validator and error formatting.
- Deliver first passing unit tests for spec integrity.

### M2: Tracked changes end-to-end

- Implement tracked change mutator.
- Add API/CLI wiring.
- Validate against golden tracked-change scenarios.

### M3: Comments end-to-end

- Implement modern threaded comments mutator.
- Add reply/resolution support and integrity checks.
- Validate against comment goldens.

### M4: Numbering end-to-end

- Implement multilevel numbering mutator.
- Add restart behavior and nested application.
- Validate against numbering goldens.

### M5: Integration hardening + docs

- Combined fixture scenarios.
- Compatibility lane/checklist baseline run.
- Publish Phase 1 user/developer docs.

## 8. Success Metrics

- At least 7 curated golden scenarios created and documented.
- >=95% pass rate for Phase 1 integration suite in CI (excluding optional Word lane).
- 100% of baseline scenarios generated via CLI from spec files.
- Zero Word "repair" prompts in the maintained golden + baseline generated fixture set during manual compatibility checks.

## 9. Risks and Mitigations

1. **Risk:** Word-specific normalization differs from direct XML mutations.
   - **Mitigation:** maintain manual goldens and periodic Word round-trip checks.
2. **Risk:** Comment/thread internals are brittle.
   - **Mitigation:** strict semantic integrity checks and targeted invalid-case tests.
3. **Risk:** Numbering edge cases explode in complexity.
   - **Mitigation:** constrain supported numbering templates in Phase 1 and document boundaries.
4. **Risk:** Cross-runtime users need stable contract for specs.
   - **Mitigation:** publish versioned spec schema and migration notes for changes.

## 10. Questions and Answers

1. Which minimum Word versions/builds are officially in compatibility scope for Phase 1? ANSWER: ignore - target up-to-date Word only
2. Do we need a JSON schema for spec interchange in addition to Python typed models? ANSWER: Phase 2
3. Should comment resolution include additional metadata beyond state (e.g., resolved-by, timestamp) in Phase 1 or Phase 2? ANSWER: No
4. What is the minimum baseline corpus size needed before declaring Phase 1 complete if edge cases emerge? ANSWER: impossible to tell at this point - keep an eye on it and update as we go
