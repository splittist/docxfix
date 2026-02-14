# Resolved Comment

*Scenario*: Document containing a paragraph with a resolved comment

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create lorem paragraph with =lorem(1,5)
* highlight "dolor sit amet"
* insert comment
* resolve that comment
* save file

*Expected visible behaviour*: a resolved comment indicator that when clicked, brings up the comment and highlights the anchor text

*Expected key XML markers*:

* in `document.xml`

  * a <w:commentRangeStart> element prior to the run containing the "dolor sit amet" text element
  * a <w:commentRangeEnd> element after the run containing the "dolor sit amet" text element
  * a run containing a <w:commentReference> element corresponding to the comment
  * a <w:commentRangeEnd> element corresponding to the reply comment

* a `comments.xml` containing:

  * a <w:comment> element with a w:id attribute whose value corresponds to that of the <w:commentRangeStart>, <w:commentRangeEnd> and <w:commentReference> elements in `document.xml`

* a `commentsExtended.xml` containing:

  * a <w15:commentEx> elements with
    * a w15:paraId attribute whose value (an 8 hex digit string) corresponds to the value of the w14:paraId attribute of the last (only) paragraph in the comment text in the <w:comment> elment in `comments.xml`
    * a w15:done attribute with the value "1", indicating the comment is resolved

* a `commentsIds.xml` containing:

  * a <w16cid:commentId> element with a w16cid:commentId attribute whose value is the value of the corresponding <w15:commentEx> w15:paraId attribute and a w16cid:durableId attribute whose value is an 8 hex digit string that does not appear elsewhere in the docx archive
