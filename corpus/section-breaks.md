*Scenario*: Two-section document demonstrating portrait → landscape transition with section-specific headers/footers.

*Word version*: Authored as a curated fixture for docxfix section validation.

*Steps taken in UI*:
1. Start from a blank document.
2. Keep section 1 portrait with a default header.
3. Insert a next-page section break.
4. Set section 2 to landscape.
5. Configure section 2 with a default header, first-page header, and default footer.

*Expected visible behaviour*:
- Page orientation switches from portrait to landscape at the section break.
- Section 2 first page shows a different header than subsequent pages.
- Footer content for section 2 is independent from section 1.

*Expected key XML markers*:
- Paragraph-level `<w:pPr><w:sectPr><w:type w:val="nextPage"/>…</w:sectPr></w:pPr>` at end of section 1.
- Body-level terminal `<w:sectPr>` for section 2 with `w:pgSz w:orient="landscape"`.
- `<w:headerReference>` and `<w:footerReference>` entries in each section mapped via `word/_rels/document.xml.rels`.
- Header/footer part overrides in `[Content_Types].xml`.

*Notes*:
- This fixture is intentionally compact for parser/validator integration checks.


*Retrieval / local generation*:
- This fixture binary is intentionally excluded from this PR.
- Generate it locally with: `python scripts/generate_section_breaks_fixture.py corpus/section-breaks.docx`.
- Then add it with: `git add corpus/section-breaks.docx`.
