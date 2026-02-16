"""Generate the section-breaks corpus fixture locally.

Usage:
    python scripts/generate_section_breaks_fixture.py [output_path]
"""

from __future__ import annotations

import sys
from pathlib import Path

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec, HeaderFooterSet, PageOrientation


def build_spec() -> DocumentSpec:
    """Build the curated multi-section fixture spec."""
    spec = DocumentSpec(seed=42)
    spec.add_paragraph("Section 1 portrait")
    spec.add_paragraph("Section 2 landscape")

    spec.sections[0].headers = HeaderFooterSet(default="Default Header S1")
    spec.add_section(
        1,
        orientation=PageOrientation.LANDSCAPE,
        headers=HeaderFooterSet(
            default="Default Header S2",
            first="First Header S2",
        ),
        footers=HeaderFooterSet(default="Default Footer S2"),
    )
    return spec


def main() -> None:
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("corpus/section-breaks.docx")
    output.parent.mkdir(parents=True, exist_ok=True)
    DocumentGenerator(build_spec()).generate(output)
    print(output)


if __name__ == "__main__":
    main()
