# docxfix

A CLI utility for creating docx fixtures with desirable characteristics for testing.

Examples include:

* tracked changes (insertions, deletions and moves)

* comments, including 'modern' comments with rich text, replies and 'resolved/done' status

* complex automatic numbering (legal-list and heading-based/styled patterns)

* highlighted text

* footnotes and endnotes

* multiple sections with different header and footer patterns

## Roadmap

A lightweight project outline is available in [`docs/project-outline.md`](docs/project-outline.md).

## Repository Layout

- `./schemas` contains the OOXML schema set used as a validation reference.
- `./corpus` contains curated `.docx` fixtures and sidecar `.md` descriptions.
- Corpus details live in [corpus/README.md](corpus/README.md).

## Features

* Built with modern Python tools: **uv**, **ruff**, **Typer**, **pytest**, **syrupy**, **lxml**, **Faker**
* Type-safe with **typing-extensions** and **types-lxml**
* XML manipulation utilities using **lxml**
* Comprehensive test suite with snapshot testing

## Installation

### Using uv (recommended)

```bash
uv pip install -e ".[dev]"
```

### Using pip

```bash
pip install -e ".[dev]"
```

## Development

### Running Tests

```bash
uv run pytest                                  # all tests
uv run pytest tests/test_generator.py          # single file
uv run pytest -k "test_heading_numbering"      # by keyword
```

### Linting and Formatting

```bash
# Check code quality
uv run ruff check src/ tests/

# Format code
uv run ruff format src/ tests/
```

### CLI Usage

```bash
# Get help
uv run docxfix --help

# Display version and info
uv run docxfix info

# Create a docx fixture
uv run docxfix create output.docx

# Create with verbose output
uv run docxfix create output.docx --verbose
```

## Visual Verification (Windows)

On Windows with Microsoft Word installed, you can export generated `.docx` files to PDF for visual inspection. This renders the document through Word's own engine, showing tracked changes, comment balloons, and numbering exactly as an end user would see them.

```bash
# Install the Windows-only extra
uv pip install -e ".[win]"

# Generate a fixture and export to PDF
uv run docxfix create scratch_out/test.docx
python scripts/export_pdf.py scratch_out/test.docx

# PDF is written to scratch_out/pdf/test.pdf
```

The script accepts individual `.docx` files or directories (defaults to `scratch_out/`). PDFs include full markup — tracked changes and comments are visible in the output.

## Project Structure

```txt
docxfix/
├── corpus/                # Golden fixtures with sidecar descriptions
├── docs/                  # Project docs and PRD
├── schemas/               # OOXML schema reference set
├── scripts/               # Utility scripts (PDF export, etc.)
├── src/
│   └── docxfix/
│       ├── __init__.py      # Package initialization
│       ├── cli.py           # CLI application (Typer)
│       ├── spec.py          # Typed spec model (DocumentSpec, Paragraph, etc.)
│       ├── generator.py     # Spec → OOXML ZIP generator
│       ├── validator.py     # Post-generation validation
│       └── xml_utils.py     # XML manipulation utilities (lxml)
├── tests/                   # pytest suite (92 tests)
└── pyproject.toml           # Project configuration
```

## Technologies

* **uv**: Fast Python package installer and resolver
* **ruff**: Extremely fast Python linter and formatter
* **Typer**: Modern CLI framework with type hints
* **pytest**: Testing framework
* **syrupy**: Snapshot testing for pytest
* **lxml**: Powerful XML processing library
* **Faker**: Deterministic fake data generation for fixtures
* **typing-extensions**: Backported and experimental type hints
* **types-lxml**: Type stubs for lxml
