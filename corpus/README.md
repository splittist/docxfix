# Corpus

This directory contains curated ".docx" fixtures and sidecar ".md" descriptions.

## File pairs

Each fixture is stored as a pair:

- "<name>.docx": the Word file used as a golden reference.
- "<name>.md": a description of how the file was created and what to verify.

## Sidecar format

Sidecar files are short, human-authored notes. The typical sections are:

- "Scenario": a one-line description of the intent.
- "Word version": the Word build used to author the file.
- "Steps taken in UI": a minimal repro checklist.
- "Expected visible behaviour": what should be seen in Word.
- "Expected key XML markers": specific XML features to validate.
- "Notes" (optional): extra implementation context.

Section titles are written as emphasized labels (for example, "*Scenario*:").
Keep wording concise and focus on the smallest set of checks needed to confirm the feature.
