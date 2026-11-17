# Verify content, structure, and appearance separately

For a newly created or modified DOCX, the repeatable entry point is:

```powershell
python scripts/docx_ops.py verify output.docx --render auto --output-dir render-output
```

For read-only inspection, check the requested findings against the source. For HTML/Markdown/text conversion, inspect the actual output's requested content, structure, and conversion diagnostics; do not require a new DOCX just to run this command.

The JSON report contains `structure`, `reopen`, `render`, and `limitations`. Use `--render never` for a deterministic check on machines without an office application. `--render word` tries only `Word.Application`; use `--render auto` for WPS through `KWPS.Application` and `--render libreoffice` when LibreOffice is the known target.

`render.status` distinguishes `skipped`, `unavailable`, `failed`, and `rendered`. In auto mode, `unavailable` means no suitable renderer was found. `failed` records a rendering failure; the error list also explains unavailable explicit selections and output conflicts. Auto fallback retains `attempts`/`errors`, including failures before a later successful export. The top-level status describes structure and reopening; also check the render status and CLI exit code before claiming rendering succeeded. A preview failure retains the exported PDF and appears in `render.pages.warning` and `limitations`.

## Three levels of evidence

1. **Content:** Requested values, replacement count, required sections, table rows, notes/comments, and selected revision view are correct. Check more than `Document.paragraphs` when content lives in other stories.
2. **Structure:** Reopen the DOCX, inspect package errors and part differences, confirm relationships/media/numbering/section changes are intended. The inspector reports advanced feature counts with the story parts where they occur. Structural success does not prove page layout.
3. **Appearance and behavior:** Open/render in the target office engine, check pages and editable behavior such as fields, hyperlinks, revisions, notes, and image wrapping. Report the renderer and any untested behavior.

```powershell
python scripts/docx_ops.py inspect output.docx
python scripts/docx_ops.py compare original.docx output.docx
```

For surgical text edits, normally only the selected story parts change. A python-docx/docxtpl save can legitimately change more; inspect rather than automatically reject every difference. The comparison checks uncompressed payload bytes, not ZIP timestamps/compression or semantic equivalence.

## Office rendering

Use the user's target application when exact output matters. WPS, Microsoft Word, and LibreOffice can paginate differently. Installed fonts, substitution, compatibility settings, and printer metrics affect layout. A `Word.Application` COM registry key alone does not prove Microsoft Word is installed or operational; compatibility applications and stale registrations exist.

The helper uses a dedicated empty work directory for every renderer attempt and publishes a PDF only after a successful export. It refuses an existing destination PDF. Choose a separate output directory for each verification run so its PDF and page previews remain identifiable.

For manual LibreOffice export, use a fresh empty output directory and a separate user profile when another instance is active:

```powershell
soffice --headless --convert-to pdf --outdir render-output output.docx
```

Use an absolute `file:///...` profile URI with `-env:UserInstallation=...` when isolation is needed. Record LibreOffice version and treat the PDF as LibreOffice's rendering, not proof of Word-identical output.

Microsoft Word automation can open a copy, update selected fields, repaginate, and call `ExportAsFixedFormat` to PDF. Disable macros and automatic external-link updates for inspection. Use a fresh dedicated application instance; close only the document/instance the task created. Do not attach to and quit the user's running application. A COM call may hang on hidden dialogs; use a bounded worker and report failure rather than repeatedly spawning applications. Do not kill all Word/WPS processes. On Windows, `--render auto` tries `KWPS.Application` before `Word.Application`, then LibreOffice; the JSON report includes `application_progid` and the renderer label so WPS output is identified.

WPS automation availability and API compatibility depend on the installed edition. Use it only after confirming the actual application and operations work. Do not label a WPS-rendered result as Word-verified.

If no renderer can be used, deliver structural checks and explicitly state that page layout has not been visually verified. Do not substitute Mammoth HTML or Pandoc PDF as evidence of the original Word layout. `verify` can rasterize a rendered PDF with Poppler (`pdfinfo` and `pdftoppm`) into page PNGs; inspect those images rather than relying on page count alone.

## Page inspection

When a PDF renderer and PDF rasterizer are available, compare before/after page count and page images at the same scale. Inspect changed pages plus neighboring pages, first/last pages, section transitions, long tables, and headers/footers. If changes affect global styles/margins/fonts, inspect all pages.

Look for clipping, overlapping text, unexpected blank pages, isolated headings, split rows, missing pictures, broken equations, changed list numbering, missing note markers, and shifted signatures. A pixel difference is a locator for review, not a universal pass/fail threshold; text edits necessarily change pixels. If only PDF export was checked, do not claim image inspection occurred.

## Completion language

Say what happened: "Updated 4 company-name occurrences; unchanged media and relationships; package inspection and reopening passed; page layout not rendered." Use stronger fidelity language only with corresponding evidence. Never equate `ZIP/XML valid` with `100% format preserved`.

For user documents, keep originals and working copies out of source-control backups unless explicitly requested. The skill repository's tests synthesize documents and discard their temporary files.
