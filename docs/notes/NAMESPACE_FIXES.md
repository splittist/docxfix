# Comment Generation Debugging - Namespace Fixes

## Problem Identified

The generated DOCX files were failing Word's validation because they were missing comprehensive namespace declarations that Word expects, even though the minimal namespaces were technically valid OOXML.

### Key Findings from Corpus Analysis

Comparing `scratch_out/step3-comments/word/comments.xml` with `corpus/comment-thread.docx`:

**Missing in Generated Files:**
- Comprehensive namespace declarations (30+ namespaces vs. 3)
- Proper `mc:Ignorable` attribute listing all optional namespaces
- Same issue affected `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, and `document.xml`

## Changes Made

### 1. Added Comprehensive Namespace Map (`src/docxfix/generator.py`)

```python
# Comprehensive Word namespace map (for compatibility)
WORD_NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    # ... 30+ namespaces matching Word's expectations
}
```

### 2. Updated XML Generation Methods

Updated the following methods to use `WORD_NAMESPACES`:
- `_create_document()` - Main document.xml
- `_create_comments()` - comments.xml
- `_create_comments_extended()` - commentsExtended.xml  
- `_create_comments_ids()` - commentsIds.xml

All now include proper `mc:Ignorable` attributes:
```python
element.set("mc:Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")
```

### 3. Created Progressive Test Script

Created `test_progressive_comments.py` with 5 minimal test cases:
1. Single comment (no reply)
2. Two separate comments
3. Comment with one reply
4. Resolved comment
5. Comment with multiple replies

## Next Steps - REQUIRES PYTHON 3.12+

**IMPORTANT:** This project requires Python 3.12+ but your system has Python 3.10.11.

### Option 1: Install Python 3.12+

1. Install Python 3.12+ from python.org
2. Run the test script:
   ```powershell
   py -3.12 -m pip install -e .
   py -3.12 test_progressive_comments.py
   ```

3. Test each generated file in Word:
   - Open `scratch_out/test1-single-comment.docx`
   - Note if Word prompts for repair
   - If repair needed, save as `test1-single-comment-fixed.docx`
   - Repeat for all test files

### Option 2: Compare XML Manually

Compare the updated generated files with corpus files:
```powershell
# After regenerating with new code
Compare-Object (Get-Content scratch_out\step3-comments\word\comments.xml) `
               (Get-Content scratch_out\corpus-comment-thread\word\comments.xml)
```

## Automated Testing Options (Windows + Word)

### Option A: PowerShell + Word COM Automation

```powershell
# Test if Word can open file without repair
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open("C:\Users\David\Code\docxfix\scratch_out\test1-single-comment.docx")
    Write-Host "✓ File opened successfully"
    $doc.Close($false)
} catch {
    Write-Host "✗ File failed to open: $_"
} finally {
    $word.Quit()
}
```

### Option B: Python + win32com

```python
import win32com.client
import os

def test_docx_in_word(filepath):
    """Test if Word can open a DOCX file without errors."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(filepath))
        print(f"✓ {filepath} opened successfully")
        doc.Close(False)
        return True
    except Exception as e:
        print(f"✗ {filepath} failed: {e}")
        return False
    finally:
        word.Quit()
```

### Option C: Office Open XML SDK (if installed)

Use Microsoft's Open XML SDK Productivity Tool to validate:
- Download: https://github.com/OfficeDev/Open-XML-SDK  
- Validate OOXML structure programmatically
- Provides detailed error reports

## Expected Outcome

With the namespace fixes, the generated files should:
1. Open in Word without repair prompts
2. Display comments correctly in the Modern Comment Experience
3. Support comment threading and resolved states
4. Match the structural patterns of the corpus files

## If Issues Persist

Additional areas to investigate:
1. **RSID attributes** - Word uses random save IDs (`w:rsidR`, `w:rsidRPr`) that we're not generating
2. **Language attributes** - `w:lang` on runs (we could add `en-US` by default)
3. **Proof attributes** - `w:noProof` flag on comment anchor runs
4. **Date formatting** - Comment date/time format might need to match Word's format exactly

These can be added progressively once we confirm the namespace fixes resolve the primary issue.
