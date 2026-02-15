# Comment Threading Bug - Status Report

## Problem Statement

Generated DOCX files contain comments with reply relationships and resolved status metadata, but **Microsoft Word does not recognize these features**:
- Comment replies appear as separate, independent comments (not threaded)
- Comments marked as resolved do not show resolved status

## Validation Evidence

**Working in SuperDoc:**
- ✅ Comments ARE threaded (replies show as nested under parent)
- ✅ Resolved comments are recognized (they don't appear, indicating SuperDoc reads the resolved flag)
- ℹ️ SuperDoc does NOT create `commentsIds.xml` entries when saving

**Not Working in Microsoft Word:**
- ❌ Comment replies appear as independent comments
- ❌ Resolved status is not recognized
- ⚠️ Word may prompt for repair on open (namespace issues previously addressed)

**Automated Validation:**
- ✅ All 61 unit/integration tests pass
- ✅ Files are valid ZIP archives
- ✅ All XML is well-formed
- ✅ Structure matches corpus golden files

## Attempts Made

### 1. Initial Implementation
**Date:** 2026-02-14  
**What:** Implemented complete comment generation with modern features
- Created `comments.xml` (comment content)
- Created `commentsExtended.xml` (threading via `paraIdParent`, resolved via `done` attribute)
- Created `commentsIds.xml` (unique identifier mappings)
- Added comment anchoring in `document.xml` (commentRangeStart/End/Reference)

**Result:** Comments appeared but were not threaded, resolved status not recognized

### 2. Comment Anchoring Fix
**Date:** 2026-02-14  
**What:** Fixed reply comment anchoring - replies need their own anchor points
- Changed from single anchor for main comment only
- To nested anchor ranges (main comment range contains reply ranges)
- Pattern: Start Main → Start Reply → Text → End Main → Ref Main → End Reply → Ref Reply

**Result:** No change in Word behavior

### 3. Namespace Compatibility Expansion
**Date:** 2026-02-15  
**What:** Added comprehensive namespace declarations matching corpus files
- Expanded from 3-6 namespaces to 30+ namespaces
- Added `mc:Ignorable` attributes listing all optional namespaces
- Applied to: `document.xml`, `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`

**Rationale:** Corpus files from Word contain extensive namespace declarations  
**Result:** Should reduce/eliminate Word repair prompts, but threading issue persists

### 4. Remove commentsIds.xml
**Date:** 2026-02-15  
**What:** Based on SuperDoc analysis, removed `commentsIds.xml` generation entirely
- Removed file creation in generator
- Removed content type entry for `commentsIds.xml`
- Removed relationship entry (adjusted IDs: numbering now rId3, styles now rId4)
- Updated all tests to not check for `commentsIds.xml`

**Rationale:** SuperDoc (which correctly threads comments) does NOT create commentsIds.xml  
**Result:** Unknown - awaiting Word verification

## Current State

### Code Changes (Ready to Commit)
- `src/docxfix/generator.py` - No longer creates `commentsIds.xml`, adjusted relationship IDs
- `tests/test_generator.py` - Removed `commentsIds.xml` assertions and test
- `tests/test_comment_integration.py` - Removed `commentsIds.xml` structure checks
- `tests/test_numbering_integration.py` - Updated expected relationship IDs

### Generated Test Files (Needs Verification)
Location: `scratch_out/`
- `step3-comments.docx` - Comment with reply
- `step5-combined.docx` - Combined features (comments, tracking, numbering)
- `test1-single-comment.docx` - Single comment, no reply
- `test2-two-comments.docx` - Two independent comments
- `test3-comment-with-reply.docx` - Main comment + 1 reply
- `test4-resolved-comment.docx` - Resolved comment
- `test5-multiple-replies.docx` - Main comment + 2 replies

### Next Verification Steps
1. Open test files in Microsoft Word
2. Check if threading now works (replies grouped under parent)
3. Check if resolved status appears (test4)
4. Report results

## Hypotheses for Why Word Doesn't Recognize Features

### Hypothesis A: Missing Required Attributes
**Theory:** Word may require additional attributes we're not generating
- `w:rsidR`, `w:rsidRDefault`, `w:rsidRPr` - Revision Save IDs
- `w:lang` - Language tags on runs
- `w:noProof` - Spell-check flags

**Evidence:**
- Corpus files contain these attributes extensively
- Our generated files have minimal attributes
- SuperDoc may be more lenient than Word

**Test:** Add RSID generation and language attributes

### Hypothesis B: Date Format Incompatibility
**Theory:** Comment date format might not match Word's expectations
**Current:** ISO 8601 format (`2026-02-14T10:30:00`)
**Word Might Want:** Different format or timezone handling

**Test:** Compare date formats in corpus vs. generated files

### Hypothesis C: Modern Comments Requires Office 365 / Word 2019+
**Theory:** Modern Comments Experience is version-dependent
- Threading might only work in newer Word versions
- Desktop vs. Web versions might differ

**Test:** Verify Word version, try opening in Word Online

### Hypothesis D: Missing or Incorrect Relationship Types
**Theory:** Relationship type URLs might be incorrect
**Current:**
- `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`
- `http://schemas.microsoft.com/office/2011/relationships/commentsExtended`

**Test:** Verify these match corpus files exactly (they should)

### Hypothesis E: Paragraph ID / Durable ID Generation
**Theory:** Our ID generation (random 8-char hex) might not match Word's algorithm
**Current:** Using random hex strings (e.g., "A1B2C3D4")
**Word Might Want:** Specific format or checksum

**Evidence:**
- SuperDoc works with our IDs
- Word might be stricter

**Test:** Analyze ID patterns in corpus files created by Word

### Hypothesis F: commentsIds.xml Was Actually Required (Despite SuperDoc)
**Theory:** Word might need `commentsIds.xml` even though SuperDoc doesn't create it
**Counter-Evidence:** SuperDoc threads comments correctly without it

**Test:** Revert commentsIds.xml removal and re-test

## Recommended Next Steps

1. **Immediate:** Verify current state in Word (post-commentsIds.xml removal)
   - Does this fix anything?
   - New error messages?

2. **If still broken:** XML Diff Analysis
   ```powershell
   # Create a comment in Word manually
   # Save as corpus-word-native.docx
   # Extract and compare XML line-by-line
   ```

3. **Add Word-specific attributes:**
   - Generate RSID values (can use random IDs matching pattern)
   - Add `w:lang="en-US"` to runs
   - Add `w:noProof="1"` to comment anchor runs

4. **Create minimal test case:**
   - Generate simplest possible file with one comment + one reply
   - Manually edit XML to exactly match Word's structure
   - Binary search which elements/attributes matter

5. **Consider Office Interop Testing:**
   - Use Word COM API to programmatically add comments
   - Capture the exact XML Word generates
   - Reverse engineer the requirements

## Files Status

### Keep/Commit
- `src/docxfix/generator.py` (modified)
- `tests/test_*.py` (modified)
- `COMMENT_THREADING_STATUS.md` (this file)
- `PROGRESS.txt` (update with latest attempt)

### Review Then Possibly Keep
- `MANUAL_VERIFICATION.md` - Useful verification guide
- `NAMESPACE_FIXES.md` - Documents namespace fix attempt
- `Test-DocxFiles.ps1` - Automated Word COM testing script
- `test_progressive_comments.py` - Generates minimal test cases
- `test_word_baseline.py` - Creates baseline for manual Word editing

### Delete (Scratch/Temporary)
- `check_rels.py` - One-off analysis script
- `scratch_out/` - All contents (generated test files)
  - Keep in .gitignore
  - Regenerate as needed for testing
- `scratch_generate.py` - If duplicate of existing functionality
- `scratch_progressive_generate.py` - If duplicate of test_progressive_comments.py

## Success Criteria

**✅ Complete Success:**
- Comment replies appear threaded (nested) in Word's review pane
- Resolved comments show with resolved indicator (checkmark/strikethrough)
- No Word repair prompts on open
- All automated tests pass

**⚠️ Partial Success:**
- Comments appear correctly but threading doesn't work
- Fall back to independent comments (acceptable if threading can't be achieved)

**❌ Failure State:**
- Comments don't appear at all
- Word crashes or corrupts file
- File requires repair on every open
