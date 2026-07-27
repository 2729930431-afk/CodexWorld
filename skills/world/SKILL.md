---
name: world
description: Control the locally running Microsoft Word desktop app and directly edit its active or named open document in place through the World MCP. Use when the user mentions World, Word, Microsoft Word, an open DOC/DOCX, the current document or selection, rewriting text, formatting, comments, tracked changes, tables, images, headers, fields, printing, exporting, macros, or any operation that should affect a document already open in Word.
---

# World

Operate the live Word application through the `world` MCP server. Prefer semantic Word tools over filesystem DOCX editing or UI automation.

## Required workflow

1. Call `word_status` before the first read or write to identify running Word, open documents, the active document, selection, save state, and track-changes state.
2. If the user identifies a document, pass its exact title or full path in every mutating call. Otherwise target the active document.
3. Read only the smallest useful scope with `word_read`.
4. Modify the live document directly. Do not create a copy, backup, temporary DOCX, or Save As result unless the user explicitly asks.
5. Leave `save_after` omitted so mutations save in place by default. For a never-saved document, do not invent a path; report that Word still needs a destination.
6. Before adding content that has a visual presentation, inspect the nearest surrounding paragraphs, tables, or other same-type objects. If the user did not specify an appearance, match the document's local visual conventions instead of using Word or tool defaults.
7. Re-read the changed range and, when formatting matters, inspect the resulting formatting after a meaningful edit. Text-only verification is not sufficient for visual changes. Use `expected_text` for precise replacements when stale-document risk exists.

Do not ask for confirmation merely because a normal in-place edit changes the document. The user explicitly opted into direct editing. Ask only when required by a higher-level safety rule or when the target document cannot be identified safely.

## Tool selection

- `word_status`: discover Word, documents, selection, and revision/save state.
- `word_read`: read the document, selection, current paragraph, or an absolute character range.
- `word_edit`: replace or insert text while preserving the rest of the document.
- `word_find_replace`: apply Word-native find/replace, including wildcards.
- `word_format`: apply styles, font properties, and paragraph formatting.
- `word_comment`: list, add, delete, or resolve comments.
- `word_track_changes`: inspect or toggle tracked changes.
- `word_review_changes`: list, accept, or reject revisions.
- `word_table`: list, insert, read/write cells, and change table structure.
- `word_image`: insert an inline image into the requested range.
- `word_document`: activate, open, save, export, print, close, or create documents. Use creation, export, print, and Save As-like behavior only when explicitly requested.
- `word_undo_redo`: undo or redo live Word actions.
- `word_com`: access nearly the complete Word COM object model when the focused tools do not expose a needed feature.

## Editing defaults

- Directly edit the requested open document.
- Save successful mutations back to the same file.
- Preserve the current track-changes setting unless the user requests a different mode.
- When the user gives no visual requirements, preserve local context: match nearby font family, size, color, paragraph spacing, alignment, borders, shading, widths, and other relevant formatting.
- Prefer the nearest structurally equivalent object as the format reference. For a new table, use a neighboring table rather than a heading or a generic built-in table style.
- Treat a style name as only one part of the appearance. Direct formatting, conditional table formatting, theme colors, widths, and cell shading must also be preserved when present.
- Use `word_table` with its default `match_context=true` for new tables. Pass `template_table_index` when a specific existing table should be the visual template.
- Prefer range edits over replacing the whole document; whole-document replacement discards local formatting.
- Never close Word or a document unless requested.
- Never accept or reject all revisions, run a macro, print, or export unless requested.
- Do not use Computer Use for document content when World MCP is available.

## Full Word object model

Use `word_com` only for capabilities not covered by focused tools. It can get/set properties, call methods, retain returned COM objects as handles, and release handles. This includes headers/footers, fields, content controls, shapes, equations, mail merge, protection, custom XML, document properties, and application commands.

Read [object-model.md](references/object-model.md) before a multi-step `word_com` sequence.
