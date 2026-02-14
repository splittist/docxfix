# docxfix

A CLI utility for creating docx fixtures with desirable characteristics for testing.

Examples include:

* tracked changes (insertions, deletions and moves)

* comments, including 'modern' comments with rich text, replies and 'resolved/done' status

* complex automatic numbering

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

* Built with modern Python tools: **uv**, **ruff**, **Typer**, **pytest**, **syrupy**, **lxml**
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
pytest tests/
```

### Linting and Formatting

```bash
# Check code quality
ruff check src/ tests/

# Format code
ruff format src/ tests/
```

### CLI Usage

```bash
# Get help
docxfix --help

# Display version and info
docxfix info

# Create a docx fixture
docxfix create output.docx

# Create with verbose output
docxfix create output.docx --verbose

# Create using a template
docxfix create output.docx --template template.docx
```

## Project Structure

```txt
docxfix/
├── corpus/                # Golden fixtures with sidecar descriptions
├── docs/                  # Project docs and PRD
├── schemas/               # OOXML schema reference set
├── src/
│   └── docxfix/
│       ├── __init__.py      # Package initialization
│       ├── cli.py           # CLI application (Typer)
│       └── xml_utils.py     # XML manipulation utilities (lxml)
├── tests/
│   ├── conftest.py          # pytest configuration
│   ├── test_cli.py          # CLI tests
│   └── test_xml_utils.py    # XML utilities tests (with syrupy)
└── pyproject.toml           # Project configuration
```

## Technologies

* **uv**: Fast Python package installer and resolver
* **ruff**: Extremely fast Python linter and formatter
* **Typer**: Modern CLI framework with type hints
* **pytest**: Testing framework
* **syrupy**: Snapshot testing for pytest
* **lxml**: Powerful XML processing library
* **typing-extensions**: Backported and experimental type hints
* **types-lxml**: Type stubs for lxml
