# Validation evidence

Environment checked on 2026-09-21: Windows, Python 3.14.7, python-docx 1.2.0, docxtpl 0.20.2, lxml 6.1.3, Jinja2 3.1.6, Pandoc 3.11, npm Mammoth 1.12.3.

## Automated checks

`python -m unittest discover -s tests -v` passed 25 tests after the verification workflow update. Tests use synthetic temporary documents and cover cross-run replacement, retained formatting and untouched parts, Chinese/XML-sensitive strings, replacement counts, no overwrite, failure outputs, protected bookmarks/revisions/fields, explicit header scope, relationships, signed-package refusal, strict templates, repeated rows, filtered inspection, feature locations, CLI exit status, explicit renderer failure, stale preview cleanup, PDF page previews, Windows code-page decoding, Windows tool discovery, actual Pandoc conversion, and actual npm Mammoth HTML/text extraction.

The helper tests were run before helper implementation and failed because the helpers were absent. Independent baseline reasoning selected the right tools but its illustrative replacement code could leave partial output and did not prevent bookmark/hyperlink boundary changes. The implemented helper deliberately refuses protected paragraphs and publishes only after validation.

The official skill-creator `quick_validate.py` reports `Skill is valid!`. This checks metadata, not document fidelity.

An independent forward-test agent loaded the skill and confirmed all four dependency probes, but its isolated command environment failed during synthetic behavior execution and produced no final result. This is recorded as a test limitation rather than counted as a passing review.

## Rendering probe

A seven-page report DOCX was exported through the installed `Word.Application` compatibility interface using `docx_ops.py verify --render auto`. The resulting PDF was checked with `pdfinfo` and rasterized with `pdftoppm` into seven page previews. PDF metadata identifies the creator as **WPS 文字**, despite the COM application name reporting Microsoft Word 12.0. The verification report now records both signals so callers do not need to trust the COM product name alone.

PDF export succeeded; closing the automation instance returned RPC unavailable (0x800706BA). Therefore this is evidence of a successful WPS export and a cleanup compatibility limitation, not a robust Word automation workflow or broad visual fidelity certification. No personal documents were used. The probe and output files are excluded from Git.

## Boundaries

The inspector is not an OOXML XSD validator. Tests do not establish arbitrary DOCX compatibility, exact Word pagination, general tracked-change editing, or lossless format conversion. Layout-sensitive user tasks still require document-specific rendering and inspection.

## Review fixes verified on 2026-11-17

`python -m unittest discover -s tests -v` passed all 47 tests. New regressions first failed against the original code, then passed after fixes: multi-section footnotes retain literal template-like data, endnotes/comments render with strict missing-value and escaping checks, renderer attempts use empty staging directories, existing PDFs survive failures, and missing output cannot be mistaken for a successful export.

Rendering tests also cover shared LibreOffice discovery, fallback diagnostics, failed-render CLI exit status, Unicode PowerShell JSON, and retained PDFs/reports when preview generation or staging cleanup fails. The PowerShell script ran with mocked COM objects; no real Office application was started for this update.

Skill metadata validation and `git diff --check` passed. Both repository and packaged requirements parsed in offline pip dry runs without installing or upgrading dependencies. An independent agent used the updated skill to extract a synthetic DOCX to UTF-8 text, confirmed its content against the source XML and unchanged source hash, and completed without Office rendering.

## Real WPS verification on 2026-11-17

The installed Codex skill rendered a synthetic two-section template containing Chinese text, XML-sensitive characters, repeated table rows, an image, headers, footers, footnotes, endnotes, and comments. All six story parts retained the literal context value `中文 A < B & 公司 {{ 6 * 7 }}`; the expression was not evaluated a second time. Source and rendered DOCX hashes remained unchanged through Office export.

Actual WPS Office 12.1.0.29690 exported a two-page A4 PDF through `KWPS.Application` on the first attempt. PDF metadata identifies the creator as `WPS 文字`; the COM name `Microsoft Word 12.0` is a compatibility label. No renderer error or cleanup warning was reported, and the pre-existing WPS processes remained running. Both page previews were opened and inspected: Chinese glyphs, table rows, image borders, section pagination, headers, footers, footnotes, and endnotes were visible without clipping, overlap, or extra blank pages. PDF text extraction independently confirmed the expected text and literal template data. Comment content was checked in DOCX XML; WPS renders the comment anchor with a colored background.

Repeating verification into the same directory exited with failure and retained the original PDF and page-preview hashes. Rendering with a missing required context value failed without creating an output DOCX. Test inputs, reports, PDFs, and previews are synthetic local artifacts under the ignored `.work` directory.

This run exposed an empty-`Producer` metadata parsing error: the parser consumed the following date line. It also exposed UTF-8 PDF metadata mixed with Windows-encoded date text. Creator and producer values now stay on their own lines, decode independently with UTF-8 preferred and a system-encoding fallback, and report empty values as null. Regression tests cover both issues. A second real export used default Windows Python with `cp936` and UTF-8 mode disabled: creator remained `WPS 文字`, producer was null, and both page-preview files were byte-identical to the first export.

This validates the installed WPS workflow; a separate Microsoft Word installation, LibreOffice, and real fallback between engines were not exercised.

## Three-agent installer verified on 2026-11-17

`python install.py` updated the actual Claude Code, Codex, and DSH skill directories and retained complete backups outside their discoverable `skills` directories. All 15 distributed files matched the repository, with only DSH's original description retained as intended. All three installed skills passed metadata validation. Running the installer again from another working directory reported all three installations unchanged.

The updated installed Codex skill also passed another real WPS export with default Windows Python encoding (`cp936`, UTF-8 mode disabled). Creator and producer metadata were correct, source hashes stayed unchanged, and the two exported page previews were byte-identical to the inspected pages.

`python -B -m unittest discover -s tests -v` completed 63 tests: 62 passed and one was skipped because this Windows account could not create symbolic links. Real source/target junction refusal was tested successfully. Installer tests cover selective installation, `CODEX_HOME`, idempotency, complete backups, DSH description preservation, cache exclusion, unreadable source directories, and restoration after a simulated publication failure. Multiline quoted DSH descriptions are refused before changing the installation; single-line quoted values and indented block descriptions are preserved.
