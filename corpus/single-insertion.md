# Single Insertion

*Scenario*: Document containing a single insertion

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create lorem paragraph with =lorem(1,5)
* turn on tracked changes
* type "single insertion " after "Lorem ipsum dolor "
* turn off tracked changes
* save file

*Expected visible behaviour*: an insertion with the text "single insertion " betwen the text "Lorem ipsum dolor " and "sit amet"

*Expected key XML markers*: in `document.xml`, a <w:ins> element wrapping a <w:r> wrapping a <w:t> containing the text "single insertion ". The <w:ins> element will have a w:author attribute
