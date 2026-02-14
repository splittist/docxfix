# Single Deletion

*Scenario*: Document containing a single deletion

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create lorem paragraph with =lorem(1,5)
* turn on tracked changes
* highlight " dolor sit amet"
* type `Delete`
* turn off tracked changes
* save file

*Expected visible behaviour*: a deletion of the text " dolor sit amet"

*Expected key XML markers*: in `document.xml`, a <w:del> element wrapping a <w:r> wrapping a <w:delText> containing the text "dolor sit amet". The <w:del> element will have a w:author attribute
