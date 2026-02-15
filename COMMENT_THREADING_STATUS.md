# Comment Threading - Resolved

## Problem Statement (Resolved)

Generated DOCX files contained comments with reply relationships and resolved status metadata, but Microsoft Word did not recognize threading or resolved status.

## Root Cause

Two issues were identified and fixed:

### 1. Missing `commentsIds.xml`

Word requires all three comment metadata files to recognize threading:
- `comments.xml` (comment content)
- `commentsExtended.xml` (threading via `paraIdParent`, resolved via `done`)
- `commentsIds.xml` (durable ID mappings)

The file was removed in an earlier attempt based on SuperDoc analysis (SuperDoc threads correctly without it), but Word requires it. The relationship type is `http://schemas.microsoft.com/office/2016/09/relationships/commentsIds`.

### 2. Comment range marker ordering for multiple replies

When a comment has 2+ replies, the ordering of `commentRangeEnd` and `commentReference` elements in `document.xml` matters.

**Broken (interleaved end/reference pairs):**
```xml
<w:commentRangeEnd w:id="0"/>     <!-- parent -->
<w:r><w:commentReference w:id="0"/></w:r>
<w:commentRangeEnd w:id="1"/>     <!-- reply 1 -->
<w:r><w:commentReference w:id="1"/></w:r>
<w:commentRangeEnd w:id="2"/>     <!-- reply 2 -->
<w:r><w:commentReference w:id="2"/></w:r>
```

**Working (grouped ends, then grouped references):**
```xml
<w:commentRangeEnd w:id="0"/>     <!-- all ends together -->
<w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="2"/>
<w:r><w:commentReference w:id="0"/></w:r>  <!-- all refs together -->
<w:r><w:commentReference w:id="1"/></w:r>
<w:r><w:commentReference w:id="2"/></w:r>
```

Note: The interleaved pattern works for a single reply (as seen in corpus golden files created by Word). It only breaks with 2+ replies.

## Verification

- Comment replies appear threaded in Word's review pane
- Resolved comments show resolved indicator
- No Word repair prompts on open
- All automated tests pass
- Tested with: single comment, comment + 1 reply, comment + 2 replies, resolved comment

## Attempts History

1. **Initial implementation** - Comments appeared but were not threaded
2. **Comment anchoring fix** - Added reply anchor ranges (no change)
3. **Namespace expansion** - Added 30+ namespace declarations matching corpus (reduced repair prompts)
4. **Remove commentsIds.xml** - Based on SuperDoc behavior (made things worse)
5. **Restore commentsIds.xml** - Threading works for single reply
6. **Fix range marker ordering** - Threading works for multiple replies
