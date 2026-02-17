"""Parse JSON/YAML fixture descriptions into DocumentSpec objects."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    HeaderFooterSet,
    NumberedParagraph,
    PageOrientation,
    Paragraph,
    SectionSpec,
    TrackedChange,
)


class SpecParseError(Exception):
    """Error raised when a fixture spec cannot be parsed.

    Attributes:
        errors: List of (field_path, reason) tuples.
    """

    def __init__(self, errors: list[tuple[str, str]]) -> None:
        self.errors = errors
        lines = [f"  {path}: {reason}" for path, reason in errors]
        super().__init__(
            f"Fixture spec has {len(errors)} error(s):\n" + "\n".join(lines)
        )


def _require(
    data: dict[str, Any],
    key: str,
    path: str,
    expected_type: type,
    errors: list[tuple[str, str]],
) -> Any | None:
    """Validate that *key* exists in *data* and has the right type."""
    field_path = f"{path}.{key}" if path else key
    if key not in data:
        errors.append((field_path, "required field is missing"))
        return None
    val = data[key]
    if not isinstance(val, expected_type):
        errors.append(
            (field_path, f"expected {expected_type.__name__}, got {type(val).__name__}")
        )
        return None
    return val


def _optional(
    data: dict[str, Any],
    key: str,
    path: str,
    expected_type: type,
    errors: list[tuple[str, str]],
    default: Any = None,
) -> Any:
    """Validate an optional field if present."""
    if key not in data:
        return default
    field_path = f"{path}.{key}" if path else key
    val = data[key]
    if not isinstance(val, expected_type):
        errors.append(
            (field_path, f"expected {expected_type.__name__}, got {type(val).__name__}")
        )
        return default
    return val


def _parse_datetime(
    value: Any, field_path: str, errors: list[tuple[str, str]]
) -> datetime | None:
    """Parse an ISO-format datetime string."""
    if value is None:
        return None
    if not isinstance(value, str):
        errors.append(
            (field_path, f"expected ISO datetime string, got {type(value).__name__}")
        )
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        errors.append((field_path, f"invalid ISO datetime format: {value!r}"))
        return None


def _parse_comment_reply(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> CommentReply | None:
    """Parse a single comment reply."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None
    text = _require(data, "text", path, str, errors)
    if text is None:
        return None
    author = _optional(data, "author", path, str, errors, "Test User")
    date_str = data.get("date")
    date = _parse_datetime(date_str, f"{path}.date", errors) if date_str else None
    return CommentReply(text=text, author=author, date=date)


def _parse_comment(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> Comment | None:
    """Parse a single comment."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None
    text = _require(data, "text", path, str, errors)
    anchor_text = _require(data, "anchor_text", path, str, errors)
    if text is None or anchor_text is None:
        return None
    author = _optional(data, "author", path, str, errors, "Test User")
    resolved = _optional(data, "resolved", path, bool, errors, False)
    date_str = data.get("date")
    date = _parse_datetime(date_str, f"{path}.date", errors) if date_str else None

    replies: list[CommentReply] = []
    raw_replies = _optional(data, "replies", path, list, errors, [])
    for i, reply_data in enumerate(raw_replies):
        reply = _parse_comment_reply(reply_data, f"{path}.replies[{i}]", errors)
        if reply is not None:
            replies.append(reply)

    return Comment(
        text=text,
        anchor_text=anchor_text,
        author=author,
        date=date,
        replies=replies,
        resolved=resolved,
    )


def _parse_tracked_change(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> TrackedChange | None:
    """Parse a single tracked change."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None

    text = _require(data, "text", path, str, errors)
    change_type_str = _require(data, "change_type", path, str, errors)

    if text is None or change_type_str is None:
        return None

    field_path = f"{path}.change_type"
    try:
        change_type = ChangeType(change_type_str)
    except ValueError:
        valid = ", ".join(f"'{ct.value}'" for ct in ChangeType)
        errors.append(
            (field_path, f"invalid change type {change_type_str!r}; expected {valid}")
        )
        return None

    author = _optional(data, "author", path, str, errors, "Test User")
    revision_id = _optional(data, "revision_id", path, int, errors, 1)
    insert_after = _optional(data, "insert_after", path, str, errors, "")
    date_str = data.get("date")
    date = _parse_datetime(date_str, f"{path}.date", errors) if date_str else None

    return TrackedChange(
        change_type=change_type,
        text=text,
        author=author,
        date=date,
        revision_id=revision_id,
        insert_after=insert_after,
    )


def _parse_numbering(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> NumberedParagraph | None:
    """Parse numbering properties."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None
    level = _optional(data, "level", path, int, errors, 0)
    numbering_id = _optional(data, "numbering_id", path, int, errors, 1)
    return NumberedParagraph(level=level, numbering_id=numbering_id)


def _parse_paragraph(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> Paragraph | None:
    """Parse a single paragraph."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None

    text = _require(data, "text", path, str, errors)
    if text is None:
        return None

    heading_level = _optional(data, "heading_level", path, int, errors, None)
    if heading_level is not None and not 1 <= heading_level <= 4:
        errors.append(
            (f"{path}.heading_level", f"must be 1-4, got {heading_level}")
        )
        heading_level = None

    # Parse tracked changes
    tracked_changes: list[TrackedChange] = []
    raw_tc = _optional(data, "tracked_changes", path, list, errors, [])
    for i, tc_data in enumerate(raw_tc):
        tc = _parse_tracked_change(tc_data, f"{path}.tracked_changes[{i}]", errors)
        if tc is not None:
            tracked_changes.append(tc)

    # Parse comments
    comments: list[Comment] = []
    raw_comments = _optional(data, "comments", path, list, errors, [])
    for i, c_data in enumerate(raw_comments):
        c = _parse_comment(c_data, f"{path}.comments[{i}]", errors)
        if c is not None:
            comments.append(c)

    # Parse numbering
    numbering = None
    if "numbering" in data:
        numbering = _parse_numbering(data["numbering"], f"{path}.numbering", errors)

    return Paragraph(
        text=text,
        tracked_changes=tracked_changes,
        comments=comments,
        numbering=numbering,
        heading_level=heading_level,
    )


def _parse_header_footer_set(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> HeaderFooterSet:
    """Parse header/footer set."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return HeaderFooterSet()
    default = _optional(data, "default", path, str, errors, None)
    first = _optional(data, "first", path, str, errors, None)
    even = _optional(data, "even", path, str, errors, None)
    return HeaderFooterSet(default=default, first=first, even=even)


def _parse_section(
    data: Any, path: str, errors: list[tuple[str, str]]
) -> SectionSpec | None:
    """Parse a single section spec."""
    if not isinstance(data, dict):
        errors.append((path, f"expected object, got {type(data).__name__}"))
        return None

    start_paragraph = _require(data, "start_paragraph", path, int, errors)
    if start_paragraph is None:
        return None

    break_type = _optional(data, "break_type", path, str, errors, "nextPage")
    restart_page_numbering = _optional(
        data, "restart_page_numbering", path, bool, errors, False
    )
    page_number_start = _optional(
        data, "page_number_start", path, int, errors, None
    )

    orientation_str = _optional(data, "orientation", path, str, errors, "portrait")
    field_path = f"{path}.orientation"
    try:
        orientation = PageOrientation(orientation_str)
    except ValueError:
        valid = ", ".join(f"'{o.value}'" for o in PageOrientation)
        errors.append(
            (field_path, f"invalid orientation {orientation_str!r}; expected {valid}")
        )
        orientation = PageOrientation.PORTRAIT

    headers = HeaderFooterSet()
    if "headers" in data:
        headers = _parse_header_footer_set(data["headers"], f"{path}.headers", errors)

    footers = HeaderFooterSet()
    if "footers" in data:
        footers = _parse_header_footer_set(data["footers"], f"{path}.footers", errors)

    return SectionSpec(
        start_paragraph=start_paragraph,
        break_type=break_type,
        orientation=orientation,
        restart_page_numbering=restart_page_numbering,
        page_number_start=page_number_start,
        headers=headers,
        footers=footers,
    )


def _parse_spec_dict(data: Any) -> DocumentSpec:
    """Parse a raw dict into a DocumentSpec, collecting all errors."""
    errors: list[tuple[str, str]] = []

    if not isinstance(data, dict):
        raise SpecParseError(
            [("$", f"expected top-level object, got {type(data).__name__}")]
        )

    title = _optional(data, "title", "$", str, errors, "Test Document")
    author = _optional(data, "author", "$", str, errors, "Test User")
    seed = _optional(data, "seed", "$", int, errors, None)

    # Parse paragraphs (required, non-empty)
    raw_paragraphs = _require(data, "paragraphs", "$", list, errors)
    paragraphs: list[Paragraph] = []
    if raw_paragraphs is not None:
        if len(raw_paragraphs) == 0:
            errors.append(("$.paragraphs", "must contain at least one paragraph"))
        for i, p_data in enumerate(raw_paragraphs):
            p = _parse_paragraph(p_data, f"$.paragraphs[{i}]", errors)
            if p is not None:
                paragraphs.append(p)

    # Parse sections (optional)
    sections: list[SectionSpec] = []
    raw_sections = _optional(data, "sections", "$", list, errors, None)
    if raw_sections is not None:
        for i, s_data in enumerate(raw_sections):
            s = _parse_section(s_data, f"$.sections[{i}]", errors)
            if s is not None:
                sections.append(s)

    if errors:
        raise SpecParseError(errors)

    spec = DocumentSpec(
        title=title,
        author=author,
        seed=seed,
        paragraphs=paragraphs,
        sections=sections if sections else [],
    )
    return spec


def parse_spec_file(path: str | Path) -> DocumentSpec:
    """Parse a JSON or YAML fixture spec file into a DocumentSpec.

    The file format is determined by extension:
    - ``.json`` → JSON
    - ``.yaml`` / ``.yml`` → YAML

    Raises:
        SpecParseError: If the spec has validation errors (with field paths).
        FileNotFoundError: If the file does not exist.
        ValueError: If the file extension is unsupported.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Spec file not found: {path}")

    text = path.read_text(encoding="utf-8")
    suffix = path.suffix.lower()

    if suffix == ".json":
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise SpecParseError(
                [("$", f"invalid JSON: {exc.args[0]}")]
            ) from exc
    elif suffix in (".yaml", ".yml"):
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise SpecParseError(
                [("$", f"invalid YAML: {exc}")]
            ) from exc
    else:
        raise ValueError(
            f"Unsupported file extension {suffix!r}; expected .json, .yaml, or .yml"
        )

    return _parse_spec_dict(data)


def parse_spec_string(text: str, *, format: str = "yaml") -> DocumentSpec:
    """Parse a JSON or YAML string into a DocumentSpec.

    Args:
        text: The spec content as a string.
        format: ``"json"`` or ``"yaml"`` (default ``"yaml"``).

    Raises:
        SpecParseError: If the spec has validation errors.
        ValueError: If *format* is unsupported.
    """
    if format == "json":
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise SpecParseError(
                [("$", f"invalid JSON: {exc.args[0]}")]
            ) from exc
    elif format == "yaml":
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise SpecParseError(
                [("$", f"invalid YAML: {exc}")]
            ) from exc
    else:
        raise ValueError(f"Unsupported format {format!r}; expected 'json' or 'yaml'")

    return _parse_spec_dict(data)
