# Styled Numbering

*Scenario*: Document containing paragraphs with styled numbering applied

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create new Multilevel List
* link first four levels of list to Heading1 to Heading4
* choose numbering properties (1, 1.1, (a), (i))
* modify Heading1 to Heading4 styles
* add text and apply various headings to paragraphs
* save file

*Expected visible behaviour*: a document with automatically numbered paragraphs with the heading styles applied

*Expected key XML markers*:

* in `document.xml`, <w:p> paragraph elements containing paragraph property <w:pPr> elements with <w:pStyle> elements with w:val attributes of "Heading1", "Heading2" etc. Note that there are NO <w:numPr> or <w:ilvl> elements.

* in `styles.xml`, <w:style> elements with w:styleId attributes corresponding to the w:pStyle values "Heading1", "Heading2" etc.

  * Each <w:style> element contains a paragraph properties <w:pPr> element with a <w:numId> entry and (except for "Heading1") a <w:ilvl> entry.
  * The <w:numId> entry corresponds to the id of the concrete <w:num> w:numId entry in `numbering.xml`
  * The <w:ilvl> entry corresponds to the w:ilvl of the abstract numbering description <w:abstractNum> with the w:abstractNumId attribute value of the concrete entry, also in `numbering.xml`

* in `numbering.xml`, a concrete <w:num> entry pointing to a <w:abstractNum> entry with <w:lvl> elements corresponding to "Heading1" to "Heading4" (etc.), and defining the numbering appearance (including the <w:lvlText> elements)

*Comment:* It's actually much easier to create complex numbering schemes by programatically manipulating the xml in `styles.xml` and `numbering.xml` than using the UI.
