# OOXML preservation and advanced edits

## Package model

A DOCX is an OPC ZIP package. Visible content is split across stories and linked parts. Editing only `word/document.xml` does not cover headers, footers, notes, or comments.

| Part | Purpose / dependencies |
|---|---|
| `[Content_Types].xml` | Default/override MIME types for parts |
| `_rels/.rels` | Package relationships, including the main document |
| `word/document.xml` | Main body, tables, section properties, drawing anchors |
| `word/_rels/document.xml.rels` | Main-story relationship IDs to media, styles, notes, etc. |
| `word/styles.xml`, `numbering.xml` | Style hierarchy, abstract numbering, numbering instances |
| `word/header*.xml`, `footer*.xml` | Per-section stories; each can have its own `.rels` |
| `word/footnotes.xml`, `endnotes.xml`, `comments.xml` | Note/comment bodies referenced by IDs |
| `word/media/*`, `embeddings/*`, `charts/*` | Binary and XML payloads; preserve unless intentionally changed |
| `word/settings.xml`, theme, font table | Field/update settings, defaults, font and theme mappings |

The helper supports conventional **Transitional OOXML** with `word/document.xml`. Strict OOXML, encrypted Office files, legacy `.doc`, macro-enabled files, and signed package editing are outside its contract. For those, use an appropriate office conversion or document-specific tooling and preserve the original.

## Conservative literal replacement

```powershell
python scripts/docx_ops.py inspect input.docx --text
python scripts/docx_ops.py replace input.docx output.docx --old 'Acme Corporation' --new 'Northwind & Co' --expected 2
python scripts/docx_ops.py compare input.docx output.docx
```

Default scope is main-document paragraphs, including tables. Explicitly add stories when intended:

```powershell
python scripts/docx_ops.py replace input.docx output.docx --old 'Acme Corporation' --new 'Northwind & Co' --expected 4 --part 'word/document.xml' --part 'word/header*.xml' --part 'word/footer*.xml'
```

The helper matches literal, case-sensitive, non-overlapping occurrences across `w:t` nodes in plain runs of a single paragraph. It applies matches from right to left, changes only affected text nodes, and preserves all unmodified part payloads. Replacement text inherits the first matched run; the remaining runs retain their properties. This is a stated formatting policy, not a unique interpretation of mixed-format replacements.

It refuses matching paragraphs with hyperlinks, bookmarks, fields, tracked edits, drawings, tabs/breaks, content controls, or other unsupported inline markup. Complex field state is tracked across paragraph boundaries. The helper is intentionally conservative: a field or drawing elsewhere in the same matched paragraph may cause refusal even if a bespoke edit could succeed.

Before publishing, it validates the whole operation, checks the expected count, writes a temporary package, and runs package diagnostics. It never overwrites the source or an existing output. Failure leaves no published DOCX. Whitespace is retained with `xml:space="preserve"`.

## A refusal is a routing decision

Read the affected paragraph and its surrounding XML, establish the intended editable range, and select a targeted operation. Do not remove protective nodes just to make the helper run. A document-specific script can edit a `w:t` inside a known hyperlink while keeping its relationship, for example, but crossing from ordinary text into a hyperlink changes semantics and needs an explicit intended outcome.

| Structure | What to preserve or decide |
|---|---|
| Hyperlinks | `r:id` targets, anchor names, link range and display text; text changes do not automatically change the target |
| Bookmarks/comments | Matching start/end IDs and ranges; moving replacement text outside a range changes its meaning |
| Fields | `begin` / `separate` / `end`, instructions, cached result; updating result text alone may be overwritten on refresh |
| Tracked changes | `w:ins`, `w:del`, `w:delText`, author/date/ID and formatting revisions; accepted/original/review views differ |
| Content controls | `w:sdtPr`, bindings, tags, placeholders, locked state; changing display text may conflict with bound custom XML |
| Numbering | `numId` -> `abstractNumId`, levels, overrides, restarts, style links; copying text/style alone does not copy list identity |
| Sections | `w:sectPr` at body end or in paragraph properties; deletion can merge sections and change headers/orientation |
| Tables | `tblGrid`, `gridSpan`, vertical merges, cell widths, row height/splitting; row/cell counts alone can be misleading |
| Floating images | `wp:anchor`, offsets, wrapping, extents, relationship ID, media and optional SVG/fallback pair |
| Equations | OMML structure and fonts; plain-text replacement or image substitution loses editability |
| Alternate content | `mc:AlternateContent` Choice/Fallback must remain consistent; do not edit both as duplicated visible text |

## Surgical ZIP editing principles

Use an XML parser with namespace-aware element access, no external entity resolution, and no recovery of malformed XML. Do not apply regex over raw XML. Preserve namespaces and extension attributes, including `mc:Ignorable`. Keep untouched parts byte-identical; serialization of a changed part may change lexical XML without changing its meaning.

Relationship IDs are scoped to their owning part. Adding an image to a header needs that header's relationships, not only the document's. When adding/removing parts, update content types and applicable relationships. When copying a body between packages, reconcile styles, numbering, relationships, notes/comments, media, and IDs; never paste raw XML and assume it is self-contained.

A ZIP can open and XML can parse while Word still repairs the document. The included diagnostics check ZIP/CRC reads, XML well-formedness, basic parts, and internal relationship targets, **not every schema constraint or referenced relationship ID**. Reopen in an office application and render when fidelity matters.

Sources: [Microsoft WordprocessingML structure](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document), [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/), and the corresponding python-docx analysis documentation. Verify advanced element ordering against the applicable schema before introducing new nodes.
