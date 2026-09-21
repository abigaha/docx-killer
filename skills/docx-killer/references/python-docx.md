# python-docx: structural editing

Use the public API for documents you create or changes whose structural scope is understood. It loads and saves a package but does not run Word's layout engine. A save may rewrite multiple XML parts; inspect the result rather than assume byte preservation.

## Read in order and identify the story

```python
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

doc = Document('input.docx')
for block in doc.iter_inner_content():
    if isinstance(block, Paragraph):
        print(block.style.name, block.text)
    elif isinstance(block, Table):
        for row in block.rows:
            print([cell.text for cell in row.cells])
```

`Document.paragraphs` excludes table-cell paragraphs, headers, footers, and many wrapper-contained paragraphs. It is not an inventory of all visible text. `iter_inner_content()` preserves top-level paragraph/table order; recurse into cells for nested tables. Merged cells may appear repeatedly through `row.cells`; deduplicate by the underlying `cell._tc` identity when mutating them. Use package inspection for notes, text boxes, revisions, and content controls.

In python-docx 1.2, `Paragraph.iter_inner_content()` yields runs and hyperlinks in order. `paragraph.runs` does not flatten hyperlink runs. `Run.text` preserves its run properties but replaces run content, potentially deleting a drawing or field marker in that run. Assign it only to a run verified to contain ordinary text.

## Create a styled document

```python
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn

doc = Document()
normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(11)
normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), 'Microsoft YaHei')
doc.add_heading('Quarterly report', level=1)
doc.add_paragraph('Prepared for the customer.')
table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
table.rows[0].cells[0].text = 'Item'
table.rows[0].cells[1].text = 'Status'
cells = table.add_row().cells
cells[0].text, cells[1].text = 'Delivery', 'Complete'
doc.sections[0].top_margin = Inches(0.8)
doc.save('new-report.docx')
```

Choose fonts available in the target renderer. Setting `font.name` alone does not reliably set the East Asian font mapping. Built-in Word style lookups use English names (`Heading 1`, `Normal`) even with a localized Word UI; custom names come from the actual document. `None` often means inherited formatting, not false or zero. Prefer styles over repeated direct formatting when editing a document's design.

## Common structural operations

| Operation | API and constraint |
|---|---|
| Insert before a paragraph | `p.insert_paragraph_before(text, style)`; adding with `doc.add_paragraph()` appends near the end, not at the current selection |
| Delete a paragraph/table | Target its element with `getparent().remove(...)`; first check section properties, bookmarks, comments, and anchors in the deleted content |
| Change page orientation | Set `section.orientation` and deliberately set/swap `page_width` and `page_height`; do not assume orientation swaps dimensions |
| Paragraph pagination | `paragraph_format.keep_with_next`, `keep_together`, `page_break_before`, `widow_control`; these are instructions to the renderer |
| Tables | Set `table.autofit = False`, grid/column widths and applicable cell widths; long text and merged cells still require page inspection |
| Table row splitting/header repeat | OOXML `w:cantSplit` / `w:tblHeader` under `w:trPr`; preserve other row properties |
| Headers/footers | `section.header`, `footer`, `first_page_header`, `even_page_header`; inspect linked sections and first/even page flags before editing |
| Inline pictures | `run.add_picture(path, width=...)`; omit height to preserve aspect ratio; inspect the specific run before replacing |
| Comments | In 1.2, `doc.add_comment(runs=[first_run, last_run], text='...', author='...', initials='...')`; anchor to actual run boundaries |

Unlinking a header can create a new blank definition, so clone intended content and dependencies when needed. Editing a linked header changes what earlier/later linked sections display. Avoid accessing every header in a way that accidentally materializes empty definitions.

### What requires OOXML or an office application

Floating image placement/wrapping, native tracked-change authoring, arbitrary footnote editing, content controls, complex numbering, equations, and field recalculation are not general high-level python-docx operations. See [OOXML](ooxml.md). Do not hand-roll native revisions by merely turning on tracking in settings; that does not retroactively record Python edits as revisions.

Tables of contents, page counts, and fields may contain stale cached text. `w:updateFields` requests an update; it is not an actual update. python-docx does not render pages or export PDF.

Sources: [text API](https://python-docx.readthedocs.io/en/latest/api/text.html), [document API](https://python-docx.readthedocs.io/en/latest/api/document.html), [tables](https://python-docx.readthedocs.io/en/latest/api/table.html), [headers](https://python-docx.readthedocs.io/en/latest/user/hdrftr.html), [comments](https://python-docx.readthedocs.io/en/latest/user/comments.html).
