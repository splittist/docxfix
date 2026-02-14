# Mixed insert delete

*Scenario*: Document containing a paragraph with a single insertion and a single deletion

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create lorem paragraph with =lorem(1,5)
* turn on tracked changes
* highlight " dolor sit amet"
* drag to before "ipsum"
* turn off tracked changes
* save file

*Expected visible behaviour*: an insertion of the text "dolor sit amet " before "ipsum" and a corresponding  deletion of the text " dolor sit amet" after "ipsum"

*Expected key XML markers*: in `document.xml`, a <w:ins> element wrapping a <w:r> wrapping a <w:t> containing the text "single insertion ", and a <w:del> element wrapping a <w:r> wrapping a <w:delText> containing the text "dolor sit amet". The <w:ins> element and the <w:del> element will each have a w:author attribute
