"""BDD row mapping example: table rows -> DocumentSpec -> .docx fixtures.

This script demonstrates the full BDD row -> spec -> docx workflow using
:func:`docxfix.bdd_mapping.map_row_to_spec`.  It can be run directly:

    uv run python examples/bdd_row_example.py

Each dict in the ``SCENARIO_TABLE`` below represents one row of a BDD
scenario outline table.  The script generates a .docx file for every row
and prints a summary.
"""

import sys
import tempfile
from pathlib import Path

from docxfix.bdd_mapping import map_row_to_spec
from docxfix.generator import DocumentGenerator
from docxfix.validator import validate_docx

# ---------------------------------------------------------------------------
# Scenario table — mimics a Gherkin "Examples:" or pytest-bdd parametrize
# ---------------------------------------------------------------------------

SCENARIO_TABLE = [
    {
        "id": "tracked-only",
        "tracked_changes": "on",
        "comment_threads": 0,
        "numbering_depth": 0,
        "use_sections": "off",
    },
    {
        "id": "comments-only",
        "tracked_changes": "off",
        "comment_threads": 2,
        "numbering_depth": 0,
        "use_sections": "off",
    },
    {
        "id": "numbered-list",
        "tracked_changes": "off",
        "comment_threads": 0,
        "numbering_depth": 3,
        "use_sections": "off",
    },
    {
        "id": "with-sections",
        "tracked_changes": "off",
        "comment_threads": 0,
        "numbering_depth": 0,
        "use_sections": "on",
    },
    {
        "id": "full-featured",
        "tracked_changes": "on",
        "comment_threads": 2,
        "numbering_depth": 2,
        "use_sections": "on",
    },
]


def main() -> int:
    success = 0
    failed = 0

    with tempfile.TemporaryDirectory() as out_dir:
        out_path = Path(out_dir)
        print(f"Generating {len(SCENARIO_TABLE)} BDD fixtures to: {out_path}\n")

        for row in SCENARIO_TABLE:
            fixture_id = row["id"]
            # Strip 'id' from the alias dict before passing to map_row_to_spec
            aliases = {k: v for k, v in row.items() if k != "id"}

            try:
                spec = map_row_to_spec(
                    aliases,
                    title=f"BDD Fixture — {fixture_id}",
                    author="Test Author",
                    seed=42,
                )
                output_file = out_path / f"{fixture_id}.docx"
                DocumentGenerator(spec).generate(str(output_file))
                validate_docx(str(output_file))
                size = output_file.stat().st_size
                print(f"  [OK] {fixture_id}  ({size:,} bytes)")
                success += 1
            except Exception as exc:
                print(f"  [FAIL] {fixture_id}: {exc}")
                failed += 1

        print(f"\nDone: {success} succeeded, {failed} failed.")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
