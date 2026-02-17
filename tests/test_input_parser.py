"""Tests for input_parser module: JSON/YAML spec parsing and validation."""

import json
import tempfile
from datetime import datetime
from pathlib import Path

import pytest

from docxfix.input_parser import SpecParseError, parse_spec_file, parse_spec_string
from docxfix.spec import ChangeType, PageOrientation

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _write_temp(content: str, suffix: str) -> Path:
    """Write content to a temp file with the given suffix and return its path."""
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=suffix, delete=False, encoding="utf-8"
    )
    tmp.write(content)
    tmp.close()
    return Path(tmp.name)


MINIMAL_YAML = """\
paragraphs:
  - text: Hello world
"""

MINIMAL_JSON = '{"paragraphs": [{"text": "Hello world"}]}'


# ===================================================================
# Happy-path: YAML parsing
# ===================================================================


class TestYAMLHappyPath:
    def test_minimal_yaml(self):
        spec = parse_spec_string(MINIMAL_YAML)
        assert len(spec.paragraphs) == 1
        assert spec.paragraphs[0].text == "Hello world"
        assert spec.title == "Test Document"
        assert spec.author == "Test User"

    def test_title_and_author(self):
        spec = parse_spec_string(
            "title: My Doc\nauthor: Alice\nparagraphs:\n  - text: hi\n"
        )
        assert spec.title == "My Doc"
        assert spec.author == "Alice"

    def test_seed(self):
        spec = parse_spec_string("seed: 42\nparagraphs:\n  - text: hi\n")
        assert spec.seed == 42

    def test_multiple_paragraphs(self):
        spec = parse_spec_string(
            "paragraphs:\n  - text: one\n  - text: two\n  - text: three\n"
        )
        assert len(spec.paragraphs) == 3
        assert [p.text for p in spec.paragraphs] == ["one", "two", "three"]


# ===================================================================
# Happy-path: JSON parsing
# ===================================================================


class TestJSONHappyPath:
    def test_minimal_json(self):
        spec = parse_spec_string(MINIMAL_JSON, format="json")
        assert len(spec.paragraphs) == 1
        assert spec.paragraphs[0].text == "Hello world"

    def test_full_json(self):
        data = {
            "title": "JSON Doc",
            "author": "Bob",
            "seed": 99,
            "paragraphs": [{"text": "para1"}, {"text": "para2"}],
        }
        spec = parse_spec_string(json.dumps(data), format="json")
        assert spec.title == "JSON Doc"
        assert spec.author == "Bob"
        assert spec.seed == 99
        assert len(spec.paragraphs) == 2


# ===================================================================
# Happy-path: Tracked changes
# ===================================================================


class TestTrackedChanges:
    def test_insertion(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Hello world"
    tracked_changes:
      - change_type: insertion
        text: " beautiful"
        insert_after: "Hello"
        author: Editor
""")
        tc = spec.paragraphs[0].tracked_changes[0]
        assert tc.change_type == ChangeType.INSERTION
        assert tc.text == " beautiful"
        assert tc.insert_after == "Hello"
        assert tc.author == "Editor"

    def test_deletion(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Remove this word please"
    tracked_changes:
      - change_type: deletion
        text: "this "
        author: Reviewer
""")
        tc = spec.paragraphs[0].tracked_changes[0]
        assert tc.change_type == ChangeType.DELETION
        assert tc.text == "this "

    def test_revision_id(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: hello
    tracked_changes:
      - change_type: insertion
        text: x
        revision_id: 42
""")
        assert spec.paragraphs[0].tracked_changes[0].revision_id == 42

    def test_default_author(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: hello
    tracked_changes:
      - change_type: insertion
        text: x
""")
        assert spec.paragraphs[0].tracked_changes[0].author == "Test User"

    def test_date_parsing(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: hello
    tracked_changes:
      - change_type: insertion
        text: x
        date: "2024-06-15T10:30:00"
""")
        tc = spec.paragraphs[0].tracked_changes[0]
        assert tc.date == datetime(2024, 6, 15, 10, 30, 0)


# ===================================================================
# Happy-path: Comments
# ===================================================================


class TestComments:
    def test_simple_comment(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "The clause is valid."
    comments:
      - text: "Is it really?"
        anchor_text: "clause"
        author: Reviewer
""")
        c = spec.paragraphs[0].comments[0]
        assert c.text == "Is it really?"
        assert c.anchor_text == "clause"
        assert c.author == "Reviewer"
        assert c.resolved is False
        assert c.replies == []

    def test_comment_with_replies(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Liability is limited."
    comments:
      - text: "Check this."
        anchor_text: "Liability"
        replies:
          - text: "Looks fine."
            author: Alice
          - text: "Agree."
            author: Bob
""")
        c = spec.paragraphs[0].comments[0]
        assert len(c.replies) == 2
        assert c.replies[0].text == "Looks fine."
        assert c.replies[0].author == "Alice"
        assert c.replies[1].author == "Bob"

    def test_resolved_comment(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Payment is net 30."
    comments:
      - text: "Approved."
        anchor_text: "net 30"
        resolved: true
""")
        assert spec.paragraphs[0].comments[0].resolved is True


# ===================================================================
# Happy-path: Numbering
# ===================================================================


class TestNumbering:
    def test_list_numbering(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "First item"
    numbering:
      level: 0
  - text: "Sub item"
    numbering:
      level: 1
""")
        assert spec.paragraphs[0].numbering is not None
        assert spec.paragraphs[0].numbering.level == 0
        assert spec.paragraphs[1].numbering.level == 1

    def test_numbering_defaults(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: hi
    numbering: {}
""")
        assert spec.paragraphs[0].numbering.level == 0
        assert spec.paragraphs[0].numbering.numbering_id == 1

    def test_heading_level(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Section Title"
    heading_level: 2
""")
        assert spec.paragraphs[0].heading_level == 2


# ===================================================================
# Happy-path: Sections
# ===================================================================


class TestSections:
    def test_section_with_headers_footers(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: "Page 1"
  - text: "Page 2"
sections:
  - start_paragraph: 0
    headers:
      default: "Header Text"
    footers:
      default: "Footer Text"
  - start_paragraph: 1
    break_type: nextPage
    orientation: landscape
""")
        assert len(spec.sections) == 2
        assert spec.sections[0].headers.default == "Header Text"
        assert spec.sections[0].footers.default == "Footer Text"
        assert spec.sections[1].orientation == PageOrientation.LANDSCAPE
        assert spec.sections[1].break_type == "nextPage"

    def test_section_page_numbering(self):
        spec = parse_spec_string("""\
paragraphs:
  - text: hi
sections:
  - start_paragraph: 0
    restart_page_numbering: true
    page_number_start: 5
""")
        assert spec.sections[0].restart_page_numbering is True
        assert spec.sections[0].page_number_start == 5


# ===================================================================
# File-based parsing
# ===================================================================


class TestFileParsing:
    def test_yaml_file(self):
        path = _write_temp(MINIMAL_YAML, ".yaml")
        try:
            spec = parse_spec_file(path)
            assert len(spec.paragraphs) == 1
        finally:
            path.unlink()

    def test_yml_extension(self):
        path = _write_temp(MINIMAL_YAML, ".yml")
        try:
            spec = parse_spec_file(path)
            assert len(spec.paragraphs) == 1
        finally:
            path.unlink()

    def test_json_file(self):
        path = _write_temp(MINIMAL_JSON, ".json")
        try:
            spec = parse_spec_file(path)
            assert len(spec.paragraphs) == 1
        finally:
            path.unlink()

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError, match="not found"):
            parse_spec_file("/nonexistent/path.yaml")

    def test_unsupported_extension(self):
        path = _write_temp("data", ".toml")
        try:
            with pytest.raises(ValueError, match="Unsupported file extension"):
                parse_spec_file(path)
        finally:
            path.unlink()


# ===================================================================
# Validation errors: top-level
# ===================================================================


class TestTopLevelErrors:
    def test_not_an_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("- item1\n- item2\n")
        assert exc_info.value.errors[0] == ("$", "expected top-level object, got list")

    def test_missing_paragraphs(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("title: hi\n")
        assert ("$.paragraphs", "required field is missing") in exc_info.value.errors

    def test_empty_paragraphs(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("paragraphs: []\n")
        errs = exc_info.value.errors
        assert ("$.paragraphs", "must contain at least one paragraph") in errs

    def test_paragraphs_wrong_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("paragraphs: not-a-list\n")
        assert any("expected list" in r for _, r in exc_info.value.errors)

    def test_title_wrong_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("title: 123\nparagraphs:\n  - text: hi\n")
        assert any("$.title" == p for p, _ in exc_info.value.errors)

    def test_seed_wrong_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string('seed: "abc"\nparagraphs:\n  - text: hi\n')
        assert any("$.seed" == p for p, _ in exc_info.value.errors)


# ===================================================================
# Validation errors: paragraphs
# ===================================================================


class TestParagraphErrors:
    def test_paragraph_not_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("paragraphs:\n  - just a string\n")
        assert any("$.paragraphs[0]" in p for p, _ in exc_info.value.errors)

    def test_paragraph_missing_text(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("paragraphs:\n  - author: x\n")
        assert (
            "$.paragraphs[0].text",
            "required field is missing",
        ) in exc_info.value.errors

    def test_heading_level_out_of_range(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string(
                "paragraphs:\n  - text: hi\n    heading_level: 5\n"
            )
        assert any("must be 1-4" in r for _, r in exc_info.value.errors)

    def test_heading_level_zero(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string(
                "paragraphs:\n  - text: hi\n    heading_level: 0\n"
            )
        assert any("must be 1-4" in r for _, r in exc_info.value.errors)


# ===================================================================
# Validation errors: tracked changes
# ===================================================================


class TestTrackedChangeErrors:
    def test_missing_change_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    tracked_changes:
      - text: x
""")
        assert any(
            "change_type" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_invalid_change_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    tracked_changes:
      - change_type: modification
        text: x
""")
        assert any(
            "change_type" in p and "invalid change type" in r
            for p, r in exc_info.value.errors
        )

    def test_missing_text(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    tracked_changes:
      - change_type: insertion
""")
        assert any(
            "tracked_changes[0].text" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_tracked_change_not_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    tracked_changes:
      - "just a string"
""")
        assert any(
            "tracked_changes[0]" in p and "expected object" in r
            for p, r in exc_info.value.errors
        )


# ===================================================================
# Validation errors: comments
# ===================================================================


class TestCommentErrors:
    def test_missing_text(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    comments:
      - anchor_text: word
""")
        assert any(
            "comments[0].text" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_missing_anchor_text(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    comments:
      - text: note
""")
        assert any(
            "comments[0].anchor_text" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_reply_missing_text(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    comments:
      - text: note
        anchor_text: hi
        replies:
          - author: Bob
""")
        assert any(
            "replies[0].text" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_reply_not_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    comments:
      - text: note
        anchor_text: hi
        replies:
          - "just text"
""")
        assert any(
            "replies[0]" in p and "expected object" in r
            for p, r in exc_info.value.errors
        )


# ===================================================================
# Validation errors: dates
# ===================================================================


class TestDateErrors:
    def test_invalid_date_format(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    tracked_changes:
      - change_type: insertion
        text: x
        date: "not-a-date"
""")
        assert any("invalid ISO datetime" in r for _, r in exc_info.value.errors)

    def test_date_wrong_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    comments:
      - text: note
        anchor_text: hi
        date: 12345
""")
        assert any("ISO datetime string" in r for _, r in exc_info.value.errors)


# ===================================================================
# Validation errors: sections
# ===================================================================


class TestSectionErrors:
    def test_missing_start_paragraph(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
sections:
  - break_type: nextPage
""")
        assert any(
            "start_paragraph" in p and "required" in r
            for p, r in exc_info.value.errors
        )

    def test_invalid_orientation(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
sections:
  - start_paragraph: 0
    orientation: diagonal
""")
        assert any("invalid orientation" in r for _, r in exc_info.value.errors)

    def test_section_not_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
sections:
  - 42
""")
        assert any(
            "sections[0]" in p and "expected object" in r
            for p, r in exc_info.value.errors
        )


# ===================================================================
# Validation errors: numbering
# ===================================================================


class TestNumberingErrors:
    def test_numbering_not_object(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    numbering: 42
""")
        assert any(
            "numbering" in p and "expected object" in r
            for p, r in exc_info.value.errors
        )

    def test_numbering_level_wrong_type(self):
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
paragraphs:
  - text: hi
    numbering:
      level: "zero"
""")
        assert any(
            "level" in p and "expected int" in r
            for p, r in exc_info.value.errors
        )


# ===================================================================
# Multiple errors collected
# ===================================================================


class TestMultipleErrors:
    def test_collects_multiple_errors(self):
        """Parser should collect all errors, not stop at the first one."""
        with pytest.raises(SpecParseError) as exc_info:
            parse_spec_string("""\
title: 123
paragraphs:
  - author: x
  - text: hi
    tracked_changes:
      - change_type: bogus
        text: y
    comments:
      - text: note
""")
        errors = exc_info.value.errors
        # title wrong type, paragraph[0] missing text, invalid change type,
        # comment missing anchor_text
        assert len(errors) >= 4
        paths = [p for p, _ in errors]
        assert "$.title" in paths
        assert "$.paragraphs[0].text" in paths

    def test_error_message_format(self):
        """SpecParseError message includes count and field paths."""
        with pytest.raises(SpecParseError, match=r"2 error\(s\)"):
            parse_spec_string("paragraphs:\n  - nope: 1\n  - nope: 2\n")


# ===================================================================
# Malformed input
# ===================================================================


class TestMalformedInput:
    def test_invalid_json(self):
        with pytest.raises(SpecParseError, match="invalid JSON"):
            parse_spec_string("{bad json", format="json")

    def test_invalid_yaml(self):
        with pytest.raises(SpecParseError, match="invalid YAML"):
            parse_spec_string(":\n  - :\n  :\n", format="yaml")

    def test_unsupported_format(self):
        with pytest.raises(ValueError, match="Unsupported format"):
            parse_spec_string("data", format="xml")


# ===================================================================
# Example fixture files (integration)
# ===================================================================


EXAMPLE_DIR = Path(__file__).parent.parent / "examples"


@pytest.mark.parametrize(
    "filename",
    [
        "01-simple.yaml",
        "02-tracked-changes.yaml",
        "03-comments.yaml",
        "04-legal-list-numbering.yaml",
        "05-heading-numbering.yaml",
        "06-sections.yaml",
        "07-combined.yaml",
        "08-deterministic.json",
    ],
)
def test_example_fixture_parses(filename: str):
    """Every shipped example fixture spec must parse without errors."""
    path = EXAMPLE_DIR / filename
    spec = parse_spec_file(path)
    assert len(spec.paragraphs) >= 1


# ===================================================================
# End-to-end: parse → generate → validate
# ===================================================================


class TestEndToEnd:
    def test_parsed_spec_generates_valid_docx(self):
        """A parsed spec should produce a valid .docx via the generator."""
        import tempfile

        from docxfix.generator import DocumentGenerator
        from docxfix.validator import validate_docx

        spec = parse_spec_string("""\
title: E2E Test
seed: 1
paragraphs:
  - text: "Clause one with a tracked change."
    tracked_changes:
      - change_type: insertion
        text: " important"
        insert_after: "one"
  - text: "Clause two with a comment."
    comments:
      - text: "Review this."
        anchor_text: "comment"
  - text: "Item A"
    numbering:
      level: 0
  - text: "Item B"
    numbering:
      level: 1
""")
        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "test.docx"
            gen = DocumentGenerator(spec)
            gen.generate(str(out))
            errors = validate_docx(str(out))
            assert not errors

    def test_example_07_generates_valid_docx(self):
        """The combined example fixture generates a valid .docx."""
        import tempfile

        from docxfix.generator import DocumentGenerator
        from docxfix.validator import validate_docx

        spec = parse_spec_file(EXAMPLE_DIR / "07-combined.yaml")
        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "combined.docx"
            gen = DocumentGenerator(spec)
            gen.generate(str(out))
            errors = validate_docx(str(out))
            assert not errors
