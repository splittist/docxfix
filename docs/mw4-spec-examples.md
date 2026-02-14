# MW-4 Baseline Spec File Examples

This document shows concrete examples of the baseline spec files referenced by **MW-4** in `docs/phase-1-prd.md`.

## Proposed directory structure

```text
fixtures/specs/phase1/
  01-single-insertion.yaml
  02-single-deletion.yaml
  03-mixed-ins-del.yaml
  04-comment-thread.yaml
  05-comment-resolved.yaml
  06-numbering-3-level.yaml
  07-numbering-restart.yaml
  08-combined-all-features.yaml
```

- One spec file per golden scenario (1:1 mapping).
- Prefix with numeric order to keep CI output stable/readable.

## Shared conventions

- `scenario_id`: stable id that matches golden file + test case name.
- `seed`: fixed integer for reproducible generated text/metadata.
- `document.paragraphs[*].id`: anchor id used by comments and tracked-change ops.
- `expected`: lightweight assertions for integration tests.

---

## Example 1: `01-single-insertion.yaml`

```yaml
version: 1
scenario_id: phase1-01-single-insertion
seed: 101
output:
  filename: phase1-01-single-insertion.docx
metadata:
  title: "Phase 1 - Single insertion"
  subject: "docxfix fixture"
authors:
  - id: a1
    name: "Alex Rivera"
    initials: "AR"
document:
  paragraphs:
    - id: p1
      text: "The Supplier shall deliver the Services within 30 days."
tracked_changes:
  - type: insertion
    paragraph_id: p1
    at_char_index: 45
    text: "commercially reasonable "
    author_id: a1
    timestamp: "2025-02-01T10:00:00Z"
expected:
  tracked_changes:
    insertions: 1
    deletions: 0
  comments:
    threads: 0
  numbering:
    numbered_paragraphs: 0
```

## Example 2: `04-comment-thread.yaml`

```yaml
version: 1
scenario_id: phase1-04-comment-thread
seed: 104
output:
  filename: phase1-04-comment-thread.docx
authors:
  - id: a1
    name: "Jordan Lee"
    initials: "JL"
  - id: a2
    name: "Taylor Kim"
    initials: "TK"
document:
  paragraphs:
    - id: p1
      text: "The Fees are payable within 15 days of invoice receipt."
comments:
  threads:
    - id: t1
      anchor:
        paragraph_id: p1
        start_char_index: 4
        end_char_index: 8
      messages:
        - id: m1
          author_id: a1
          text: "Should this be 30 days?"
          timestamp: "2025-02-01T10:00:00Z"
        - id: m2
          reply_to: m1
          author_id: a2
          text: "Keep 15 for enterprise template."
          timestamp: "2025-02-01T10:02:00Z"
      resolved: false
expected:
  tracked_changes:
    insertions: 0
    deletions: 0
  comments:
    threads: 1
    replies: 1
    resolved_threads: 0
```

## Example 3: `08-combined-all-features.yaml`

```yaml
version: 1
scenario_id: phase1-08-combined-all-features
seed: 108
output:
  filename: phase1-08-combined-all-features.docx
authors:
  - id: a1
    name: "Casey Morgan"
    initials: "CM"
  - id: a2
    name: "Riley Shah"
    initials: "RS"
document:
  paragraphs:
    - id: p1
      text: "Services"
      numbering:
        list_id: legal-main
        level: 0
    - id: p2
      text: "Implementation services are provided remotely."
      numbering:
        list_id: legal-main
        level: 1
    - id: p3
      text: "Travel expenses require pre-approval."
      numbering:
        list_id: legal-main
        level: 1
numbering:
  definitions:
    - id: legal-main
      style: legal-decimal-paren
      levels:
        - level: 0
          format: "%1."
        - level: 1
          format: "%1.%2"
        - level: 2
          format: "(%3)"
tracked_changes:
  - type: deletion
    paragraph_id: p2
    start_char_index: 37
    end_char_index: 45
    author_id: a1
    timestamp: "2025-02-01T11:00:00Z"
  - type: insertion
    paragraph_id: p2
    at_char_index: 37
    text: "on-site or remotely"
    author_id: a2
    timestamp: "2025-02-01T11:05:00Z"
comments:
  threads:
    - id: t1
      anchor:
        paragraph_id: p3
        start_char_index: 0
        end_char_index: 15
      messages:
        - id: m1
          author_id: a2
          text: "Define approval SLA in schedule."
          timestamp: "2025-02-01T11:10:00Z"
      resolved: true
expected:
  tracked_changes:
    insertions: 1
    deletions: 1
  comments:
    threads: 1
    resolved_threads: 1
  numbering:
    numbered_paragraphs: 3
```

## Notes for implementation

- Exact field names can be adjusted when the typed spec model is finalized; preserve scenario shape and 1:1 golden mapping.
- Keep each baseline spec intentionally small and single-purpose (except the combined scenario).
- Add one integration test per file that asserts the `expected` block.
