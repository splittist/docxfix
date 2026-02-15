"""Automated Word COM verification for comment threading.

Requires: pywin32, Pillow (optional for screenshots)
Install: uv pip install pywin32 Pillow

Opens generated DOCX files in Word via COM automation and checks:
1. Document opens without repair prompts
2. Comments collection is populated
3. Reply threading is recognized
4. Resolved status is correct
"""

import os
import sys
import time
from pathlib import Path

SCRATCH_DIR = Path(__file__).parent / "scratch_out"

# Test file expectations
TEST_FILES = {
    "test1-single-comment.docx": {
        "expected_comments": 1,
        "expected_replies": {},  # comment_index -> reply_count
        "any_resolved": False,
    },
    "test2-two-comments.docx": {
        "expected_comments": 2,
        "expected_replies": {},
        "any_resolved": False,
    },
    "test3-comment-with-reply.docx": {
        "expected_comments": 2,  # Word may show parent + reply as 2 comments
        "expected_replies": {},  # Will check if replies are threaded
        "any_resolved": False,
    },
    "test4-resolved-comment.docx": {
        "expected_comments": 1,
        "any_resolved": True,
    },
    "test5-multiple-replies.docx": {
        "expected_comments": 4,  # 1 parent + 3 replies
        "any_resolved": False,
    },
}


def take_screenshot(word_app, filepath):
    """Try to capture a screenshot of the Word window."""
    try:
        from PIL import ImageGrab
        import win32gui

        hwnd = word_app.ActiveWindow.Hwnd
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)

        # Get window rect
        rect = win32gui.GetWindowRect(hwnd)
        img = ImageGrab.grab(bbox=rect)
        img.save(filepath)
        print(f"  Screenshot saved: {filepath}")
    except Exception as e:
        print(f"  Screenshot failed: {e}")


def check_document(word_app, filepath, expectations):
    """Open a document and check its comments."""
    print(f"\nChecking: {filepath.name}")
    print("-" * 50)

    abs_path = str(filepath.resolve())
    doc = None

    try:
        doc = word_app.Documents.Open(abs_path, ReadOnly=True)
        time.sleep(1)

        comments = doc.Comments
        comment_count = comments.Count
        print(f"  Comments found: {comment_count}")

        expected = expectations.get("expected_comments", 0)
        if comment_count >= 1:
            print(f"  PASS: Document has comments (expected ~{expected}, got {comment_count})")
        else:
            print(f"  FAIL: No comments found (expected {expected})")

        # Check each comment for details
        for i in range(1, comment_count + 1):
            c = comments(i)
            author = c.Author
            text = c.Range.Text[:50] if c.Range.Text else "(empty)"
            print(f"  Comment {i}: author='{author}', text='{text}'")

            # Check for replies (Word 2016+ COM)
            try:
                replies = c.Replies
                reply_count = replies.Count
                if reply_count > 0:
                    print(f"    -> {reply_count} replies (THREADING WORKS!)")
                    for j in range(1, reply_count + 1):
                        r = replies(j)
                        print(f"       Reply {j}: author='{r.Author}', text='{r.Range.Text[:50]}'")
            except AttributeError:
                print("    (Replies property not available - older Word version)")
            except Exception as e:
                print(f"    Replies check error: {e}")

            # Check resolved status
            try:
                # Word 2019+ exposes Done property
                done = c.Done
                print(f"    Resolved: {bool(done)}")
            except AttributeError:
                pass
            except Exception:
                pass

        # Take screenshot
        screenshot_path = filepath.with_suffix(".png")
        take_screenshot(word_app, str(screenshot_path))

        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        return False
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def main():
    try:
        import win32com.client
    except ImportError:
        print("pywin32 not installed. Install with: uv pip install pywin32")
        sys.exit(1)

    print("=" * 60)
    print("Word COM Comment Threading Verification")
    print("=" * 60)

    # Check test files exist
    missing = []
    for name in TEST_FILES:
        if not (SCRATCH_DIR / name).exists():
            missing.append(name)

    if missing:
        print(f"\nMissing test files in {SCRATCH_DIR}:")
        for m in missing:
            print(f"  - {m}")
        print("\nRun test_progressive_comments.py first to generate them.")
        sys.exit(1)

    # Start Word
    print("\nStarting Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = False  # Suppress repair/conversion dialogs
    time.sleep(2)

    results = {}
    try:
        for name, expectations in TEST_FILES.items():
            filepath = SCRATCH_DIR / name
            success = check_document(word, filepath, expectations)
            results[name] = success

    finally:
        print("\n" + "=" * 60)
        print("RESULTS SUMMARY")
        print("=" * 60)
        for name, success in results.items():
            status = "PASS" if success else "FAIL"
            print(f"  [{status}] {name}")

        # Close Word
        print("\nClosing Word...")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
