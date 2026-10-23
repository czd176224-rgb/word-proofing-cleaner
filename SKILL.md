---
name: word-proofing-cleaner
description: Remove Microsoft Word red/blue proofing underlines, spellcheck/grammar marks, and stale proofing error tags from .docx files while preserving the visible document text. Use when the user asks to clear red lines, blue lines, wavy underlines, spelling/grammar underlines, or proofreading marks in a Word document without changing the original content.
---

# Word Proofing Cleaner

## Purpose

Use this skill for `.docx` files where Word shows red or blue underlines caused by spelling, grammar, or proofing metadata, and the user wants the visible content preserved.

The bundled script removes existing `w:proofErr` markers, adds Word settings to hide spelling and grammar errors, and marks existing runs as `w:noProof` so Word does not immediately re-run proofing on the same text.

This suppresses proofing for existing text; it does not fix spelling or grammar mistakes. Use it for reviewed documents when hiding these marks is the requested outcome. New edits and editor-specific checking may introduce new marks. Supports `.docx`, not `.doc`, `.docm`, encrypted documents or PowerPoint.

Requires Python 3.10+ and `lxml` (`python -m pip install -r requirements.txt`). Processing stays local; the script does not upload documents or call an AI API.

## Workflow

1. Work on a copy unless the user explicitly asks to overwrite the original.
2. Run `scripts/clean_word_proofing.py` with the input `.docx` and output path.
3. Use `--verify-text` to compare visible text between input and output.
4. Report the output file path and whether visible text was unchanged.
5. If available, render or open-check the output document. If rendering tools are unavailable, state that structural validation passed but visual rendering was skipped.

## Commands

Create a cleaned copy:

```bash
python scripts/clean_word_proofing.py input.docx output-clean.docx --verify-text
```

Overwrite in place only when requested:

```bash
python scripts/clean_word_proofing.py input.docx --in-place --verify-text
```

## What It Changes

- Removes `w:proofErr` elements from Word XML parts.
- Enables `w:hideSpellingErrors` and `w:hideGrammaticalErrors` in the existing `word/settings.xml`, when present.
- Sets `w:proofState` spelling and grammar to `clean`.
- Enables `w:noProof` in existing run properties and default run properties, and creates run properties for existing text runs that have none.

## What It Must Not Change

- Do not edit visible document text.
- Do not accept or reject tracked changes unless separately requested.
- Do not remove comments, formatting, images, tables, headers, or footers.
- Do not overwrite the source document unless the user explicitly asks.

## Notes

- Red/blue proofing lines are a Word display feature, so PNG/PDF renderers may not show them even when Word does. Structural checks are more important for this specific issue.
- If the visible text comparison fails, do not deliver the output as final; inspect the changed parts first.
- The script validates text before replacing the destination, including in-place operations. A locked Word document can block the final replacement; close it and retry.
