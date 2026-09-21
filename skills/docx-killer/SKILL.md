---
name: docx-killer
description: Use when creating, inspecting, editing, templating, converting, or troubleshooting Word DOCX files, especially when preserving formatting, tables, headers, images, comments, revisions, or document structure matters. Covers python-docx, docxtpl, Pandoc, npm Mammoth, and targeted OOXML edits.
---

# DOCX Killer

Preserve the user's document intent. Choose tools by the requested operation, not by trying to run all four. A DOCX is an OOXML package; semantic conversion is not lossless editing, and valid XML is not proof of correct pagination.

## Start with the document

1. Establish the input, requested change, output, and required fidelity. Inspect the actual document before asking questions already answered by its contents. Preserve the source and use a new output by default.
2. Run `python scripts/doctor.py` from this skill directory when the environment is unknown. Python and npm Mammoth are separate installations. Pandoc can exist outside the current PATH. See [Windows setup](references/windows.md).
3. Run `python scripts/docx_ops.py inspect input.docx --text` for story/paragraph locations, structures, relationships, and package diagnostics. Text reports may contain private document content; keep them local. The inspector is limited to conventional Transitional DOCX, not full schema validation.
4. Choose the narrowest suitable operation below. When a preservation helper refuses a protected structure, inspect that structure and use a targeted edit or an office engine. Do not bypass refusal by flattening the paragraph.

## Choose the tool

| Need | Route |
|---|---|
| Create a document; edit styles, sections, tables, inline images, comments | [python-docx](references/python-docx.md) |
| Replace literal text while retaining unaffected package parts | `docx_ops.py replace`; [OOXML preservation](references/ooxml.md) |
| Populate an existing template; repeated rows, conditions, rich text, images | `docx_ops.py render`; [docxtpl](references/docxtpl.md) |
| Markdown/HTML/DOCX conversion, reference styles, citations | [Pandoc](references/pandoc.md) |
| Semantic HTML or raw text extraction | `mammoth_extract.cjs`; [Mammoth](references/mammoth.md) |
| Fields, revisions, floating objects, text boxes, numbering, exact layout | [OOXML preservation](references/ooxml.md) and [verification](references/verification.md) |

## Reliable defaults

- `replace` defaults to `word/document.xml` only, including its tables. Headers/footers/notes/comments require explicit `--part` selection. Require an independently counted `--expected`; new text inherits the first matched run's formatting. Runs can split words invisibly.
- Avoid `paragraph.text = ...` and `cell.text = ...` when preserving existing inline content. These operations rebuild content. Public APIs do not cover every story or OOXML feature.
- Render trusted templates with strict undefined values and autoescaping. JSON data is not template source. Plain text goes into ordinary variables, not arbitrary XML/RichText.
- Pandoc `--reference-doc` supplies styles and document properties; it does not preserve the original body as a template. Mammoth HTML is semantic, not a visual rendering, and is not sanitized.
- Work on a new DOCX and compare parts with `docx_ops.py compare before.docx after.docx`. Review unexpected media, relationship, numbering, style, and section changes.
- For layout-sensitive work, open/export with the user's target office engine and inspect pages. State explicitly when only structural checks were possible.

## Deliver

Return the output path, substantive edits, actual verification performed, and remaining limitations. Never describe arbitrary DOCX editing as universally lossless. Advanced edits are supported through document-specific OOXML/office workflows, not by pretending the conservative text helper handles every structure.
