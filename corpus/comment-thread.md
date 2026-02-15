# Comment Thread

*Scenario*: Document containing a paragraph with a comment and a reply to that comment

*Word version*: Microsoft® Word for Microsoft 365 MSO (Version 2601 Build 16.0.19628.20166) 64-bit

*Steps taken in UI*:

* create lorem paragraph with =lorem(1,5)
* highlight "dolor sit amet"
* insert comment
* reply to that comment
* save file

*Expected visible behaviour*: an insertion of a comment anchored to the text "dolor sit amet" and a reply comment below that comment displayed as a thread

*Expected key XML markers*:

* in `document.xml`

  * two <w:commentRangeStart> elements (one for each of the comment and the reply comment) prior to the run containing the "dolor sit amet" text element
  * a <w:commentRangeEnd> element corresponding to the original comment after the run containing the "dolor sit amet" text element
  * a run containing a <w:commentReference> element corresponding to the original comment
  * a <w:commentRangeEnd> element corresponding to the reply comment
  * a run containing a <w:commentReference> element corresponding to the reply comment

  **Note on multi-reply ordering:** This golden file has a single reply, so the end/reference elements are interleaved (end0, ref0, end1, ref1). For 2+ replies, Word requires grouping all <w:commentRangeEnd> elements together before all <w:commentReference> runs (end0, end1, end2, ref0, ref1, ref2). See COMMENT_THREADING_STATUS.md for details.

* a `comments.xml` containing:

  * two <w:comment> elements, one for each of the original comment and the reply comment
  * each <w:comment> element has a w:id attribute whose value corresponds to that of the corresponding <w:commentRangeStart>, <w:commentRangeEnd> and <w:commentReference> elements in `document.xml`

* a `commentsExtended.xml` containing:

  * two <w15:commentEx> elements, one for each of the original comment and the reply comment
  * each <w15:commentEx> element has a w15:paraId attribute whose value (an 8 hex digit string) corresponds to the value of the w14:paraId attribute of the last (only) paragraph in the comment text in the <w:comment> element in `comments.xml`
  * each <w15:commentEx> element contains a w15:done attribute with the value "0", indicating the comment is not resolved
  * the <w15:commentEx> element corresponding to the reply comment has a w15:paraIdParent attribute whose value is the value of the w15:paraId attribute of the original (parent) comment

* a `commentsIds.xml` containing:

  * two <w16cid:commentId> elements, one for each of the original comment and the reply comment
  * each <w16cid:commentId> element contains
    * a w16cid:commentId attribute whose value is the value of the corresponding <w15:commentEx> w15:paraId attribute
    * a w16cid:durableId attribute whose value is an 8 hex digit string that does not appear elsewhere in the docx archive
