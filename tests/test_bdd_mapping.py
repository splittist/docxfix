"""Tests for bdd_mapping module: BDD row -> DocumentSpec mapping."""

from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

import pytest

from docxfix.bdd_mapping import (
    VALID_ALIASES,
    BDDMappingError,
    map_row_to_spec,
)
from docxfix.generator import DocumentGenerator
from docxfix.spec import ChangeType


# ===========================================================================
# VALID_ALIASES
# ===========================================================================


class TestValidAliases:
    def test_valid_aliases_has_expected_keys(self):
        assert set(VALID_ALIASES) == {
            "tracked_changes",
            "comment_threads",
            "numbering_depth",
            "use_sections",
        }

    def test_valid_aliases_values_are_strings(self):
        for key, description in VALID_ALIASES.items():
            assert isinstance(description, str), f"{key!r} description should be str"
            assert len(description) > 0


# ===========================================================================
# Happy paths: empty / default row
# ===========================================================================


class TestEmptyRow:
    def test_empty_row_returns_spec(self):
        spec = map_row_to_spec({})
        assert spec is not None

    def test_empty_row_has_one_paragraph(self):
        spec = map_row_to_spec({})
        assert len(spec.paragraphs) == 1
        assert spec.paragraphs[0].text == "BDD fixture document."

    def test_empty_row_no_tracked_changes(self):
        spec = map_row_to_spec({})
        assert spec.paragraphs[0].tracked_changes == []

    def test_empty_row_no_comments(self):
        spec = map_row_to_spec({})
        assert spec.paragraphs[0].comments == []

    def test_default_title_and_author(self):
        spec = map_row_to_spec({})
        assert spec.title == "BDD Fixture"
        assert spec.author == "Test User"

    def test_custom_title_author_seed(self):
        spec = map_row_to_spec({}, title="My Doc", author="Alice", seed=99)
        assert spec.title == "My Doc"
        assert spec.author == "Alice"
        assert spec.seed == 99


# ===========================================================================
# Happy paths: tracked_changes
# ===========================================================================


class TestTrackedChanges:
    def test_tracked_changes_true_adds_paragraph(self):
        spec = map_row_to_spec({"tracked_changes": True})
        # At least one paragraph with tracked changes
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert len(tc_paragraphs) >= 1

    def test_tracked_changes_has_insertion_and_deletion(self):
        spec = map_row_to_spec({"tracked_changes": True})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        types = {tc.change_type for p in tc_paragraphs for tc in p.tracked_changes}
        assert ChangeType.INSERTION in types
        assert ChangeType.DELETION in types

    def test_tracked_changes_false_no_tracked_paragraphs(self):
        spec = map_row_to_spec({"tracked_changes": False})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert tc_paragraphs == []

    def test_tracked_changes_on_string(self):
        spec = map_row_to_spec({"tracked_changes": "on"})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert len(tc_paragraphs) >= 1

    def test_tracked_changes_off_string(self):
        spec = map_row_to_spec({"tracked_changes": "off"})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert tc_paragraphs == []

    def test_tracked_changes_true_string(self):
        spec = map_row_to_spec({"tracked_changes": "true"})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert len(tc_paragraphs) >= 1

    def test_tracked_changes_false_string(self):
        spec = map_row_to_spec({"tracked_changes": "false"})
        tc_paragraphs = [p for p in spec.paragraphs if p.tracked_changes]
        assert tc_paragraphs == []


# ===========================================================================
# Happy paths: comment_threads
# ===========================================================================


class TestCommentThreads:
    def test_zero_threads_no_comment_paragraphs(self):
        spec = map_row_to_spec({"comment_threads": 0})
        comment_paragraphs = [p for p in spec.paragraphs if p.comments]
        assert comment_paragraphs == []

    def test_one_thread(self):
        spec = map_row_to_spec({"comment_threads": 1})
        comment_paragraphs = [p for p in spec.paragraphs if p.comments]
        assert len(comment_paragraphs) == 1

    def test_three_threads(self):
        spec = map_row_to_spec({"comment_threads": 3})
        comment_paragraphs = [p for p in spec.paragraphs if p.comments]
        assert len(comment_paragraphs) == 3

    def test_each_thread_has_one_reply(self):
        spec = map_row_to_spec({"comment_threads": 2})
        for p in spec.paragraphs:
            for comment in p.comments:
                assert len(comment.replies) == 1

    def test_comment_threads_as_string(self):
        spec = map_row_to_spec({"comment_threads": "2"})
        comment_paragraphs = [p for p in spec.paragraphs if p.comments]
        assert len(comment_paragraphs) == 2

    def test_comment_anchor_text_matches_paragraph(self):
        spec = map_row_to_spec({"comment_threads": 1})
        p = next(p for p in spec.paragraphs if p.comments)
        assert p.comments[0].anchor_text == p.text


# ===========================================================================
# Happy paths: numbering_depth
# ===========================================================================


class TestNumberingDepth:
    def test_zero_depth_no_numbered_paragraphs(self):
        spec = map_row_to_spec({"numbering_depth": 0})
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert numbered == []

    def test_depth_one(self):
        spec = map_row_to_spec({"numbering_depth": 1})
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 1
        assert numbered[0].numbering.level == 0

    def test_depth_three(self):
        spec = map_row_to_spec({"numbering_depth": 3})
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 3
        levels = [p.numbering.level for p in numbered]
        assert levels == [0, 1, 2]

    def test_depth_four(self):
        spec = map_row_to_spec({"numbering_depth": 4})
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 4

    def test_numbering_depth_as_string(self):
        spec = map_row_to_spec({"numbering_depth": "3"})
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 3


# ===========================================================================
# Happy paths: use_sections
# ===========================================================================


class TestUseSections:
    def test_use_sections_false_default_section_only(self):
        spec = map_row_to_spec({"use_sections": False})
        # DocumentSpec always adds section at 0; no extra sections
        assert all(s.start_paragraph == 0 for s in spec.sections)

    def test_use_sections_true_adds_second_section(self):
        spec = map_row_to_spec({"use_sections": True})
        assert any(s.start_paragraph == 1 for s in spec.sections)

    def test_use_sections_true_section_has_header(self):
        spec = map_row_to_spec({"use_sections": True})
        sect = next(s for s in spec.sections if s.start_paragraph == 1)
        assert sect.headers.default is not None
        assert len(sect.headers.default) > 0

    def test_use_sections_on_string(self):
        spec = map_row_to_spec({"use_sections": "on"})
        assert any(s.start_paragraph == 1 for s in spec.sections)

    def test_use_sections_off_string(self):
        spec = map_row_to_spec({"use_sections": "off"})
        assert all(s.start_paragraph == 0 for s in spec.sections)

    def test_use_sections_true_ensures_two_paragraphs(self):
        # With no other aliases, empty row would give 1 paragraph; use_sections
        # should ensure there are at least 2.
        spec = map_row_to_spec({"use_sections": True})
        assert len(spec.paragraphs) >= 2


# ===========================================================================
# Happy paths: combined aliases
# ===========================================================================


class TestCombinedAliases:
    def test_all_aliases_on(self):
        row = {
            "tracked_changes": True,
            "comment_threads": 2,
            "numbering_depth": 3,
            "use_sections": True,
        }
        spec = map_row_to_spec(row)
        assert any(p.tracked_changes for p in spec.paragraphs)
        comment_paragraphs = [p for p in spec.paragraphs if p.comments]
        assert len(comment_paragraphs) == 2
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 3
        assert any(s.start_paragraph == 1 for s in spec.sections)

    def test_tracked_changes_and_comments(self):
        row = {"tracked_changes": "on", "comment_threads": 1}
        spec = map_row_to_spec(row)
        assert any(p.tracked_changes for p in spec.paragraphs)
        assert any(p.comments for p in spec.paragraphs)

    def test_numbering_and_sections(self):
        row = {"numbering_depth": 2, "use_sections": True}
        spec = map_row_to_spec(row)
        numbered = [p for p in spec.paragraphs if p.numbering]
        assert len(numbered) == 2
        assert any(s.start_paragraph == 1 for s in spec.sections)


# ===========================================================================
# Error cases: unknown aliases
# ===========================================================================


class TestUnknownAliases:
    def test_unknown_alias_raises(self):
        with pytest.raises(BDDMappingError, match="Unknown alias"):
            map_row_to_spec({"unknown_key": "value"})

    def test_multiple_unknown_aliases_raises(self):
        with pytest.raises(BDDMappingError, match="Unknown alias"):
            map_row_to_spec({"foo": 1, "bar": 2})

    def test_error_message_mentions_valid_aliases(self):
        with pytest.raises(BDDMappingError) as exc_info:
            map_row_to_spec({"bad": True})
        msg = str(exc_info.value)
        for alias in VALID_ALIASES:
            assert alias in msg


# ===========================================================================
# Error cases: invalid values
# ===========================================================================


class TestInvalidValues:
    def test_tracked_changes_invalid_string(self):
        with pytest.raises(BDDMappingError, match="tracked_changes"):
            map_row_to_spec({"tracked_changes": "maybe"})

    def test_tracked_changes_invalid_type(self):
        with pytest.raises(BDDMappingError, match="tracked_changes"):
            map_row_to_spec({"tracked_changes": 42})

    def test_comment_threads_negative(self):
        with pytest.raises(BDDMappingError, match="comment_threads"):
            map_row_to_spec({"comment_threads": -1})

    def test_comment_threads_float_string(self):
        with pytest.raises(BDDMappingError, match="comment_threads"):
            map_row_to_spec({"comment_threads": "1.5"})

    def test_comment_threads_bool_raises(self):
        with pytest.raises(BDDMappingError, match="comment_threads"):
            map_row_to_spec({"comment_threads": True})

    def test_numbering_depth_too_large(self):
        with pytest.raises(BDDMappingError, match="numbering_depth"):
            map_row_to_spec({"numbering_depth": 5})

    def test_numbering_depth_negative(self):
        with pytest.raises(BDDMappingError, match="numbering_depth"):
            map_row_to_spec({"numbering_depth": -1})

    def test_numbering_depth_bool_raises(self):
        with pytest.raises(BDDMappingError, match="numbering_depth"):
            map_row_to_spec({"numbering_depth": True})

    def test_use_sections_invalid_value(self):
        with pytest.raises(BDDMappingError, match="use_sections"):
            map_row_to_spec({"use_sections": "maybe"})


# ===========================================================================
# End-to-end: row -> spec -> valid .docx
# ===========================================================================


class TestEndToEnd:
    def _generate(self, spec) -> Path:
        tmpdir = tempfile.mkdtemp()
        out = Path(tmpdir) / "test.docx"
        DocumentGenerator(spec).generate(out)
        return out

    def test_empty_row_generates_valid_docx(self):
        spec = map_row_to_spec({})
        out = self._generate(spec)
        assert zipfile.is_zipfile(out)

    def test_all_aliases_generates_valid_docx(self):
        row = {
            "tracked_changes": True,
            "comment_threads": 2,
            "numbering_depth": 3,
            "use_sections": True,
        }
        spec = map_row_to_spec(row, seed=42)
        out = self._generate(spec)
        assert zipfile.is_zipfile(out)

    def test_deterministic_with_seed(self):
        row = {"tracked_changes": True, "comment_threads": 1}
        spec1 = map_row_to_spec(row, seed=7)
        spec2 = map_row_to_spec(row, seed=7)
        out1 = self._generate(spec1)
        out2 = self._generate(spec2)
        assert out1.read_bytes() == out2.read_bytes()

    def test_row_to_spec_to_docx_full_flow(self):
        """Demonstrates the full BDD row -> spec -> .docx flow."""
        # Simulates a BDD table row (e.g. from pytest-bdd parametrize)
        bdd_row = {
            "tracked_changes": "on",
            "comment_threads": "1",
            "numbering_depth": "2",
            "use_sections": "off",
        }
        spec = map_row_to_spec(bdd_row, title="Contract Fixture", seed=123)
        out = self._generate(spec)
        with zipfile.ZipFile(out, "r") as z:
            assert "word/document.xml" in z.namelist()
