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
