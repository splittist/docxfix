# Manual Comment Verification Guide

## What Was Fixed

The generator was creating reply comment metadata and linking them via `commentsExtended.xml`,
but **not anchoring the replies in the document**. Replies need their own comment ranges at the
same location as the main comment.

### Before Fix
Only the main comment had anchor points in document.xml:
```xml
<w:commentRangeStart w:id="0"/>
<w:t>ANCHOR</w:t>
<w:commentRangeEnd w:id="0"/>
<w:commentReference w:id="0"/>
```

### After Fix
Both main comment and replies are anchored (nested ranges):
```xml
<w:commentRangeStart w:id="0"/>       <!-- Main comment start -->
<w:commentRangeStart w:id="1"/>       <!-- Reply start -->
<w:t>ANCHOR</w:t>
<w:commentRangeEnd w:id="0"/>         <!-- Main comment end -->
<w:commentReference w:id="0"/>
<w:commentRangeEnd w:id="1"/>         <!-- Reply end -->
<w:commentReference w:id="1"/>
```

## Files to Verify

Open each file in Microsoft Word and verify the expected behavior:

### 1. test1-single-comment.docx
**Expected:**
- ✓ Single comment appears on word "ANCHOR"
- ✓ No replies
- ✓ Not resolved

### 2. test2-two-comments.docx
**Expected:**
- ✓ Two separate comments (not threaded)
- ✓ First comment on "ANCHOR" by "Reviewer One"
- ✓ Second comment on "MARKER" by "Reviewer Two"
- ✓ Both not resolved

### 3. test3-comment-with-reply.docx ⭐
**Expected:**
- ✓ Main comment: "Main comment." by "Reviewer One"
- ✓ Reply comment: "Reply to main comment." by "Reviewer Two" 
- ✓ Reply should appear THREADED under main comment (indented/nested)
- ✓ Not resolved

**How to check threading:**
- Click on the comment - you should see a conversation thread
- The reply should be visually grouped with the main comment
- May show "1 reply" or expand/collapse indicator

### 4. test4-resolved-comment.docx ⭐
**Expected:**
- ✓ Single comment on "ANCHOR"
- ✓ Comment should show as **RESOLVED**
- ✓ May have checkmark or different color/styling

**How to check resolved status:**
- Look for resolved indicator (✓ checkmark, strikethrough, or muted color)
- Or check comment properties/details

### 5. test5-multiple-replies.docx ⭐
**Expected:**
- ✓ Main comment: "Main comment with multiple replies." by "Reviewer One"
- ✓ First reply: "First reply." by "Reviewer Two"
- ✓ Second reply: "Second reply." by "Reviewer Three"
- ✓ All three should be threaded together
- ✓ Not resolved

### 6. step3-comments.docx
**Expected:**
- ✓ Main comment with reply, threaded
- ✓ Both from "Test User"

### 7. step5-combined.docx
**Expected:**
- ✓ Comment with reply (threaded)
- ✓ Tracked changes (insertions/deletions)
- ✓ Numbered list
- ✓ All features working together

## What to Report

If comments still don't appear correctly:

### Scenario A: Threading doesn't work
If replies appear as separate comments instead of threaded:
- Report: "Replies appear as independent comments, not threaded"
- Save the file as `test3-actual-behavior.docx` and share

### Scenario B: Resolved status doesn't show
If comment in test4 doesn't show as resolved:
- Report: "Comment appears but not marked as resolved"
- Check if Word version supports Modern Comments (requires Office 2016+)

### Scenario C: Comments don't appear at all
If some comments are missing:
- Report: "Only X of Y comments appear"
- Note which comment IDs are visible

## Testing Notes

- Modern Comments Experience requires Office 2016 or later
- Some features may look different depending on Word version
- Desktop Word vs. Word Online may render differently
- Make sure Review pane is visible: Review tab → Show Comments

## Success Criteria

✅ **Fully Working:** All comments appear, replies are threaded, resolved status shows
⚠️ **Partial:** Comments appear but threading/resolved doesn't work
❌ **Broken:** Comments missing or Word requires repair
