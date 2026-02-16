# Sections

*Scenario*: Document containing three sections including one in landscape

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create title page
* insert section break with new page
* create lorem text with =lorem(10,10)
* select lorem text, format as mulilevel list, then indent various paragraphs to provide structure
* insert section break with new page
* change orientation of new section to landscape
* format headers and footers, none linked to previous
  * title page with blank headers and footers
  * section two has header text and page number (start at 1) in footer
  * section three has header text and page number (start at 1) in footer
* save file

*Expected visible behaviour*: a four-page document with: "Title Page" on the first page, with no header or footer text; two pages of numbered lorem text, with "Header for Section 2" in the header and a 1-based page number in the footer; one landscape page with the text "Section 3, which is landscape", a header with "Header for section 3", and a page number in the footer.

*Expected key XML markers*:

* in `document.xml`

  * three <w:sectPr> elements
    * the first within the paragraph properties of the last paragraph of the title page (first section)
    * the second within the paragraph properties of the last paragraph of the body (second secton)
    * the last at the top level (i.e., not within any paragraph) covering the last section
  * the <w:sectPr> element for the title page and the body define a portrait page (<w:pgSz> with a w:w attribute smaller than the w:h attribute)
  * the <w:sectPr> element for the last section defines a landscape page (<w:pgSz> with a w:w attribute greater than the w:h attribute)
  * the <w:sectPr> element for the body and the last section have a <w:headerReference> and a <w:footerReference> with w:type "default"
  * the <w:sectPr> element for the body and the last section have a <w:pgNumType> with w:start of "1"

* `header1.xml` and `header2.xml` each containing the text of the header for the body and last sectino respectively (as linked through rId reference and _rels)

* `footer1.xml` and `footer2.xml` each containing quite a lot of boilerplate around using standard docpart, the key part of which is a run containng a <w:instrText> element with "PAGE \\* MERGEFORMAT" (again linked through rId references)
