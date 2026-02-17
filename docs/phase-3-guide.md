# Phase 3: BDD Fixture Generation Guide

This guide covers Phase 3 features: generating `.docx` fixtures from external spec files (JSON/YAML) via CLI, designed for BDD test automation workflows.

## Overview

Phase 3 enables non-Python users to generate fixtures without writing Python code:
- **External spec format:** Define fixtures in JSON or YAML
- **CLI workflows:** `create --spec` and `batch` commands
- **BDD integration:** Use generated fixtures in any test framework

## Quick Start

### 1. Create a Spec File

Create a YAML file describing your document (e.g., `my-fixture.yaml`):

```yaml
title: Contract Review Document
author: Legal Team
seed: 42  # Optional: for deterministic output
paragraphs:
  - text: This is a simple contract clause.
  - text: This clause has tracked changes text.
    tracked_changes:
      - change_type: insertion
        text: tracked changes
        author: Editor
  - text: This clause has a comment on specific text.
    comments:
      - text: Please review this section
        anchor_text: specific text
        author: Reviewer
```

### 2. Generate the Fixture

```bash
uv run docxfix create contract.docx --spec my-fixture.yaml
```

The generated `contract.docx` can now be used in your test suite.

## Spec File Format

### Supported Formats

- **YAML:** `.yaml` or `.yml` extensions
- **JSON:** `.json` extension

### Complete Spec Reference

```yaml
# Document metadata (all optional)
title: "Document Title"
author: "Author Name"
seed: 42  # For deterministic output in CI

# Paragraphs (required, must have at least one)
paragraphs:
  # Simple paragraph
  - text: "Plain paragraph text"
  
  # Paragraph with tracked changes
  - text: "Base text with insertion here"
    tracked_changes:
      - change_type: insertion
        text: "insertion"
        author: "Editor"
        date: "2024-01-15T10:30:00"
        revision_id: 1
        insert_after: "with "
      
      - change_type: deletion
        text: "deleted text"
        author: "Editor"
        date: "2024-01-15T10:31:00"
        revision_id: 2
  
  # Paragraph with comments
  - text: "Text with anchor for comment"
    comments:
      - text: "Main comment text"
        anchor_text: "anchor"
        author: "Reviewer"
        date: "2024-01-15T14:00:00"
        resolved: false
        replies:
          - text: "Reply to the comment"
            author: "Author"
            date: "2024-01-15T15:00:00"
  
  # Numbered paragraph (legal-list style)
  - text: "First numbered item"
    numbering:
      level: 0
  
  - text: "Nested numbered item"
    numbering:
      level: 1
  
  # Heading-based numbering
  - text: "Section heading"
    heading_level: 1
  
  - text: "Subsection heading"
    heading_level: 2

# Sections (optional, for multi-section documents)
sections:
  - start_paragraph: 0
    orientation: portrait
    header_text: "Section 1 Header"
    footer_text: "Page {page}"
  
  - start_paragraph: 3
    orientation: landscape
    header_text: "Section 2 Header"
    footer_text: "Page {page} of {total}"
```

### Field Validation

The parser validates:
- **Required fields:** `paragraphs` with at least one paragraph, each with `text`
- **Type correctness:** Strings, integers, booleans, dates as specified
- **Enum values:** `change_type` (insertion/deletion), `orientation` (portrait/landscape)
- **Date format:** ISO 8601 format (`YYYY-MM-DDTHH:MM:SS`)
- **Numeric ranges:** `heading_level` (1-4), `numbering.level` (0-8)

Errors include field paths (e.g., `$.paragraphs[0].text`) and specific reasons.

## CLI Commands

### `create --spec`

Generate a single fixture from a spec file.

```bash
uv run docxfix create OUTPUT --spec SPEC_FILE [OPTIONS]

Arguments:
  OUTPUT          Output path for the generated docx file

Options:
  --spec, -s      Path to spec file (JSON/YAML)
  --validate      Validate generated document (default: on)
  --no-validate   Skip validation
  --verbose, -v   Enable verbose output
```

**Examples:**

```bash
# Basic usage
uv run docxfix create output.docx --spec examples/01-simple.yaml

# With verbose output
uv run docxfix create output.docx --spec my-spec.yaml --verbose

# Skip validation (faster, but not recommended)
uv run docxfix create output.docx --spec my-spec.yaml --no-validate
```

### `batch`

Generate multiple fixtures from a batch manifest.

```bash
uv run docxfix batch --manifest MANIFEST [OPTIONS]

Options:
  --manifest, -m  Path to batch manifest file (YAML)
  --out-dir, -o   Output directory (default: fixtures)
  --validate      Validate generated documents (default: on)
  --no-validate   Skip validation
  --verbose, -v   Enable verbose output
```

**Manifest Format:**

```yaml
fixtures:
  - id: fixture-identifier
    spec: path/to/spec.yaml
    output: output-filename.docx
  
  - id: another-fixture
    spec: path/to/another-spec.json
    output: another-output.docx
```

**Notes:**
- Spec paths are relative to the manifest file directory
- Output files are written to `--out-dir`
- `id` is used for error reporting
- Batch exits with code 1 if any fixture fails
- Failed fixtures are listed at the end

**Example:**

```bash
uv run docxfix batch --manifest test-fixtures.yaml --out-dir ./fixtures --verbose
```

**Sample Output:**

```
✓ simple-doc
✓ complex-doc
✓ tracked-changes-doc
============================================================
Batch generation complete:
  Total: 3
  Success: 3
  Failed: 0
```

## Example Fixtures

Eight example spec files are included in `examples/`:

1. **`01-simple.yaml`** — Minimal document with plain paragraphs
2. **`02-tracked-changes.yaml`** — Insertions and deletions
3. **`03-comments.yaml`** — Threaded comments with replies and resolved state
4. **`04-legal-list-numbering.yaml`** — Multilevel legal-list numbering
5. **`05-heading-numbering.yaml`** — Heading-based styled numbering
6. **`06-sections.yaml`** — Multiple sections with headers, footers, and page numbering
7. **`07-combined.yaml`** — All features combined
8. **`08-deterministic.json`** — Seeded deterministic output (JSON format)

Try them:

```bash
uv run docxfix create test.docx --spec examples/07-combined.yaml
```

## BDD Integration Examples

### Python with pytest-bdd

```python
# features/steps/document_steps.py
from pytest_bdd import scenarios, given, when, then
from pathlib import Path
import subprocess

scenarios('../document_validation.feature')

@given('a document with tracked changes')
def document_with_tracked_changes(tmp_path):
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text("""
title: Test Document
paragraphs:
  - text: "Text with insertion here"
    tracked_changes:
      - change_type: insertion
        text: "insertion"
        author: "Editor"
""")
    
    output = tmp_path / "test.docx"
    subprocess.run([
        "uv", "run", "docxfix", "create",
        str(output), "--spec", str(spec_file)
    ], check=True)
    
    return output

@then('the document should contain tracked changes')
def verify_tracked_changes(document_with_tracked_changes):
    # Use your document analysis tool here
    assert document_with_tracked_changes.exists()
```

### JavaScript/TypeScript with Cucumber

```javascript
// features/step_definitions/document_steps.js
const { Given, When, Then } = require('@cucumber/cucumber');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

Given('a document with comments', function () {
  const specFile = path.join(this.tmpDir, 'spec.yaml');
  fs.writeFileSync(specFile, `
title: Test Document
paragraphs:
  - text: "Text with anchor"
    comments:
      - text: "Test comment"
        anchor_text: "anchor"
        author: "Reviewer"
`);
  
  this.outputFile = path.join(this.tmpDir, 'test.docx');
  execSync(`uv run docxfix create ${this.outputFile} --spec ${specFile}`);
});

Then('the document should contain comments', function () {
  // Use your document analysis tool here
  const exists = fs.existsSync(this.outputFile);
  expect(exists).toBe(true);
});
```

### Scenario Tables

For BDD scenarios with table-driven tests, consider creating a helper script to map scenario table rows to fixture specs. See the Phase 3 PRD section on "BDD row mapping helper" for future tooling in this area.

## Error Handling

### Spec Parsing Errors

If a spec file has errors, the CLI reports all validation issues:

```
✗ Spec parsing error: Validation failed with 3 errors:
  - $.paragraphs[0]: missing required field 'text'
  - $.paragraphs[1].tracked_changes[0]: 'change_type' must be 'insertion' or 'deletion', got 'modification'
  - $.paragraphs[2].comments[0]: missing required field 'anchor_text'
```

### Batch Errors

In batch mode, individual fixture failures don't stop processing:

```
✓ fixture-1
✗ Fixture 'fixture-2': Spec parsing error: missing required field 'paragraphs'
✓ fixture-3
============================================================
Batch generation complete:
  Total: 3
  Success: 2
  Failed: 1

Failed fixtures:
  - fixture-2: Spec error: missing required field 'paragraphs'
```

## Deterministic Output

For CI reproducibility, specify a `seed` in your spec:

```yaml
title: Deterministic Document
seed: 42
paragraphs:
  - text: "This will generate identically on every run"
```

With the same seed, generated documents are byte-identical across runs and platforms.

## Validation

By default, all generated documents are validated:
- **ZIP structure:** Proper `.docx` archive format
- **XML well-formedness:** All XML parts parse correctly
- **Semantic checks:** ID uniqueness, relationship integrity, content type coverage

Disable validation with `--no-validate` for faster generation (not recommended for CI).

## Next Steps

- **Phase 3 M3.3:** BDD row mapping helper for scenario table integration
- **Phase 3 M3.4:** Extended examples and quick-start templates

## Limitations

- Comment anchoring requires exact text match (no fuzzy matching)
- One comment per paragraph currently (multiple comments create sequential anchors)
- Section breaks must reference valid paragraph indices
- Date fields must be ISO 8601 format

## Related Documentation

- [Phase 1 Guide](phase-1-guide.md) — Python API and architecture
- [Phase 3 PRD](future-phases-prd.md) — Full Phase 3 specification
- [Project Outline](project-outline.md) — Roadmap and direction
