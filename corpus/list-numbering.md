# List Numbering

*Scenario*: Document containing five paragraphs with legal-style list numbering of varous levels

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create five lorem paragraphs with =lorem(5,5)
* highlight all paragraphs
* format as a Multilevel List with 1, 1.1, 1.1.1 (legal) style
* indent the second paragraph once
* indent the fourth paragraph once
* indent the fifth paragraph twice
* save file

*Expected visible behaviour*: a five paragraph file, with paragraphs numbered 1., 1.1., 2., 2.1, and 2.1.1

*Expected key XML markers*:

* in `document.xml`

  * five <w:p> elements, each containing:
    * a <w:pPr> element with a <w:pStyle> element with w:val attribute "ListParagraph"
    * a <w:numPr> element with:
      * a <w:numId> element with a w:val attribute of "1"
      * a <w:ilvl> element with a w:val attribute corresponding to the zero-based numbering depth, from "0" (top level) to "2" for the fith paragraph ("2.1.1")

* a `numbering.xml` containing:
  * a <w:num> element with a w:numId attribute of "1" (corresponding to the numId of the paragraphs), and containing a <w:abstractNumId> element with a w:val attribute of "0"
  * a <w:abstractNum> element with a w:abstractNumId attribute with a value of "0" (corresonding to the w:abstractNumId value of the (concrete) <w:num> element) containing:
    * a <w:multiLevalType> elmeent with a w:val attribute of "multileval"
    * <w:lvl> elements, each with a w:ilvl attribute with values corresponding to the various w:ilvl element values, and the various formatting to be applied, including <w:lvlText> elements whose w:val attributes contain strings indicating how the numbering should be displayed, e.g. "%1.", "%1.%2.%3." etc.

* a `styles.xml` containing a <w:style> element with w:styleId attribute "ListParagraph"
