"""BDD row mapping helper: map table row keys to DocumentSpec fixtures.

This module provides :func:`map_row_to_spec`, which converts a BDD-style table
row (a plain :class:`dict`) into a :class:`~docxfix.spec.DocumentSpec` ready
for generation.

Supported aliases
-----------------
- ``tracked_changes`` — bool or ``'on'``/``'off'``: adds a paragraph with an
  insertion and a deletion tracked change.
- ``comment_threads`` — int (≥ 0): adds *N* paragraphs each anchored to a
  comment thread with one reply.
- ``numbering_depth`` — int (0–4): adds one legal-list paragraph per depth
  level (levels 0 … depth-1).
- ``use_sections`` — bool or ``'on'``/``'off'``: adds a second document
  section with a default header.

Example
-------
.. code-block:: python

    from docxfix.bdd_mapping import map_row_to_spec
    from docxfix.generator import DocumentGenerator

    # Typical BDD table row (e.g. from pytest-bdd or behave)
    row = {
        "tracked_changes": "on",
        "comment_threads": 2,
        "numbering_depth": 3,
        "use_sections": "off",
    }

    spec = map_row_to_spec(row, title="Contract v1", seed=42)
    DocumentGenerator(spec).generate("contract_v1.docx")
"""

from __future__ import annotations

from typing import Any

from docxfix.spec import (
    ChangeType,
    Comment,
    CommentReply,
    DocumentSpec,
    HeaderFooterSet,
    NumberedParagraph,
    Paragraph,
    SectionSpec,
    TrackedChange,
)

#: Supported alias keys and their human-readable descriptions.
VALID_ALIASES: dict[str, str] = {
    "tracked_changes": (
        "Enable tracked changes (bool or 'on'/'off'). "
        "Adds a paragraph with an insertion and a deletion."
    ),
    "comment_threads": (
        "Number of comment threads to include (int >= 0). "
        "Each thread contains one reply."
    ),
    "numbering_depth": (
        "Depth of legal-list numbering (int 0-4). "
        "Adds one paragraph per numbering level."
    ),
    "use_sections": (
        "Enable a second document section (bool or 'on'/'off'). "
        "Adds a section with a default header."
    ),
}


class BDDMappingError(Exception):
    """Raised when a BDD row contains unknown aliases or invalid values."""


def _parse_bool(value: Any, alias: str) -> bool:
    """Coerce *value* to bool; accept ``True``/``False``, ``'on'``/``'off'``,
    ``'true'``/``'false'``, ``'yes'``/``'no'``."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        if value.lower() in ("on", "true", "yes"):
            return True
        if value.lower() in ("off", "false", "no"):
            return False
    raise BDDMappingError(
        f"Alias {alias!r}: expected bool or 'on'/'off', got {value!r}"
    )


def _parse_non_neg_int(value: Any, alias: str) -> int:
    """Coerce *value* to a non-negative int."""
    if isinstance(value, bool):
        raise BDDMappingError(
            f"Alias {alias!r}: expected integer, got bool"
        )
    if isinstance(value, int):
        result = value
    elif isinstance(value, str):
        try:
            result = int(value)
        except ValueError:
            raise BDDMappingError(
                f"Alias {alias!r}: expected integer, got {value!r}"
            ) from None
    else:
        raise BDDMappingError(
            f"Alias {alias!r}: expected integer, got {value!r}"
        )
    if result < 0:
        raise BDDMappingError(
            f"Alias {alias!r}: must be >= 0, got {result}"
        )
    return result


def map_row_to_spec(
    row: dict[str, Any],
    *,
    title: str = "BDD Fixture",
    author: str = "Test User",
    seed: int | None = None,
) -> DocumentSpec:
    """Map a BDD table row to a :class:`~docxfix.spec.DocumentSpec`.

    Args:
        row: Dict of alias keys to values (typically one row from a BDD table).
        title: Document title (default ``'BDD Fixture'``).
        author: Document author used for tracked changes and comments.
        seed: Optional RNG seed for deterministic output.

    Returns:
        A :class:`~docxfix.spec.DocumentSpec` ready for
        :class:`~docxfix.generator.DocumentGenerator`.

    Raises:
        BDDMappingError: If the row contains unknown aliases or invalid values.
    """
    unknown = set(row) - set(VALID_ALIASES)
    if unknown:
        valid = ", ".join(sorted(VALID_ALIASES))
        raise BDDMappingError(
            f"Unknown alias(es): {', '.join(sorted(unknown))}. "
            f"Valid aliases: {valid}"
        )

    paragraphs: list[Paragraph] = []
    extra_sections: list[SectionSpec] = []

    # --- tracked_changes ---
    if "tracked_changes" in row:
        use_tracked = _parse_bool(row["tracked_changes"], "tracked_changes")
        if use_tracked:
            paragraphs.append(
                Paragraph(
                    text="Sample contract clause.",
                    tracked_changes=[
                        TrackedChange(
                            change_type=ChangeType.INSERTION,
                            text="INSERTED ",
                            author=author,
                        ),
                        TrackedChange(
                            change_type=ChangeType.DELETION,
                            text="Sample",
                            author=author,
                        ),
                    ],
                )
            )

    # --- comment_threads ---
    if "comment_threads" in row:
        n_threads = _parse_non_neg_int(row["comment_threads"], "comment_threads")
        for i in range(n_threads):
            target = f"Contract clause {i + 1}."
            paragraphs.append(
                Paragraph(
                    text=target,
                    comments=[
                        Comment(
                            text=f"Review thread {i + 1}",
                            anchor_text=target,
                            author=author,
                            replies=[
                                CommentReply(text="Acknowledged.", author=author)
                            ],
                        )
                    ],
                )
            )

    # --- numbering_depth ---
    if "numbering_depth" in row:
        depth = _parse_non_neg_int(row["numbering_depth"], "numbering_depth")
        if depth > 4:
            raise BDDMappingError(
                f"Alias 'numbering_depth': must be 0-4, got {depth}"
            )
        for level in range(depth):
            paragraphs.append(
                Paragraph(
                    text=f"Numbered item at level {level + 1}.",
                    numbering=NumberedParagraph(level=level),
                )
            )

    # Ensure at least one paragraph
    if not paragraphs:
        paragraphs.append(Paragraph(text="BDD fixture document."))

    # --- use_sections ---
    if "use_sections" in row:
        use_sects = _parse_bool(row["use_sections"], "use_sections")
        if use_sects:
            # Ensure there are at least 2 paragraphs for the second section
            if len(paragraphs) < 2:
                paragraphs.append(Paragraph(text="Second section content."))
            extra_sections.append(
                SectionSpec(
                    start_paragraph=1,
                    headers=HeaderFooterSet(default=f"{title} — Section 2"),
                )
            )

    return DocumentSpec(
        title=title,
        author=author,
        seed=seed,
        paragraphs=paragraphs,
        sections=extra_sections,
    )
