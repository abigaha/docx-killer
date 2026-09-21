# Pandoc: semantic conversion and generated DOCX

Pandoc parses a document into its intermediate representation and writes a new document. It is suitable for conversions, structured authoring, citations, and style-based generation. It is not a surgical editor and cannot promise a DOCX round trip preserving every Word feature.

Use the executable reported by `doctor.py`; the examples assume `pandoc` is on PATH. Every output path should be a new file unless replacement was intended.

```powershell
pandoc input.docx -f docx -t markdown --extract-media=extracted-media -o extracted.md
pandoc input.docx -f docx+styles -t json -o structure.json
pandoc report.md -f markdown -t docx --reference-doc=reference.docx -o report.docx
pandoc article.html -f html -t docx -o article.docx
pandoc report.md --citeproc --bibliography=references.bib --csl=style.csl -o cited-report.docx
```

`--extract-media` extracts images to disk and may write files; choose a fresh directory. Resources/filters may read files or execute code: use known inputs and filters appropriate to the task. No filter is required for ordinary DOCX conversion.

## Reference document versus template

Create the default reference using binary-safe output:

```powershell
pandoc -o reference.docx --print-default-data-file reference.docx
```

Edit the styles in Word/WPS/LibreOffice, then use `--reference-doc=reference.docx`. Pandoc uses its styles and document properties, including margins, page size, headers, and footers; **it ignores its body content**. Use docxtpl when the body itself is a designed template that must remain.

For custom paragraph/character styles, use Pandoc-native Div/Span `custom-style` attributes and define the styles in the reference document. Read `-f docx+styles` into an output format that retains attributes, such as JSON or Pandoc Markdown. An attribute surviving the AST does not prove equivalent rendering or preservation of arbitrary Word properties.

## Revisions and comments

Pandoc's DOCX reader defaults to `--track-changes=accept`. That selects the accepted text and ignores comments. Use `--track-changes=reject` to select the rejected text, or `--track-changes=all` to expose insertions/deletions/comments as annotated AST spans for analysis. These options affect the reader; they do not guarantee re-creation of native Word review markup on output.

Do not silently run a revision-bearing contract through the default conversion when the user needs original review history. Inspect the package and establish the intended revision view.

## Fidelity boundaries

Detailed table geometry, floating shapes, text boxes, section boundaries, fields, linked objects, equation formatting, and complex numbering may be lost or transformed. Generated DOCX may retain useful structural content while differing in page breaks and appearance. A Markdown or HTML round trip cannot restore information discarded on the first conversion.

Pandoc PDF output typically uses a separate PDF engine. Converting DOCX to PDF through Pandoc is reconstruction, not Word rendering. To check the original DOCX layout, use the target office application's PDF export, as described in [verification](verification.md).

Source: [Pandoc manual](https://pandoc.org/MANUAL.html), sections `--reference-doc`, `--track-changes`, `--extract-media`, `styles` extension, and Custom styles. Verified installed executable: 3.11.
