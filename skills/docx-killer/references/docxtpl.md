# docxtpl: render an existing designed template

Use when the template defines the layout and variables/conditions/repeated regions define the changing content. It is not a substitute for editing arbitrary existing prose.

## Strict, escaped rendering

```powershell
python scripts/docx_ops.py render invoice-template.docx context.json invoice.docx
```

The helper uses `Environment(undefined=StrictUndefined, autoescape=True)` and `render(..., autoescape=True)`. A missing nested property fails rather than silently producing an empty string; autoescaping protects ordinary text in XML. Neither validates `null`, empty values, arithmetic correctness, or business requirements.

The CLI renders existing footnotes, endnotes, and comments once per part, including parts shared by multiple story relationships. Check rendered values separately in the body, headers/footers, notes, and comments; a body-only text check cannot confirm the other stories.

Example context:

```json
{
  "customer": {"name": "A < B & Partners"},
  "items": [{"label": "Consulting", "quantity": 2, "amount": "300.00"}]
}
```

Use trusted template files. A normal Jinja environment is not a sandbox for malicious templates. The JSON CLI intentionally handles ordinary data; use the Python API for `RichText`, `InlineImage`, or subdocuments.

## Place tags in Word correctly

| Template expression | Placement |
|---|---|
| `{{ customer.name }}` | Keep the complete tag within one run; ordinary text inherits the template formatting |
| `{%p if show_details %}` / `{%p endif %}` | Separate control paragraphs surrounding the visible paragraphs |
| `{%tr for item in items %}` / `{%tr endfor %}` | Separate control rows surrounding one or more detail rows |
| `{%tc ... %}` | Dedicated table-column control locations; inspect the resulting table grid |
| `{{r rich_value }}` | A run placeholder for a `RichText` object |
| `{% cellbg color %}` | At the beginning of the cell; color is a hex value without `#` |

Control paragraphs/rows are removed during rendering. Do not put both opening and closing structural tags into the same control paragraph/row. Keep spaces inside delimiters as documented. Word may split visually continuous template tags into runs after styling edits; inspect the XML when a template fails.

For a repeating table, make a control row containing `{%tr for item in items %}`, a styled detail row containing `{{ item.label }}`, and a final control row containing `{%tr endfor %}`. Access dictionary keys carefully: `item.items` can resolve a dict method; use `item['items']` for colliding names.

## Rich text, images, and subdocuments

```python
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm
from jinja2 import Environment, StrictUndefined

tpl = DocxTemplate('template.docx')
name = RichText()
name.add('Important', bold=True, color='B00020')
context = {'rich_value': name, 'photo': InlineImage(tpl, 'photo.png', width=Mm(35))}
tpl.render(context, jinja_env=Environment(undefined=StrictUndefined, autoescape=True), autoescape=True)
tpl.save('generated.docx')
```

`{{r ... }}` replaces character formatting with the RichText formatting; set required font/style explicitly. Combine segments with `.add()` rather than placing multiple `{{r ... }}` placeholders in the same run. Jinja filters do not operate on RichText in the same way as plain strings; transform the input before constructing RichText.

Use `InlineImage` for body images. For header/footer pictures, the documented workflow is a dummy image in the template followed by media replacement; dynamic addition is limited. Check dimensions/aspect ratio. `replace_pic`, `replace_media`, and `replace_zipname` have different matching semantics, so inspect the specific method before choosing one.

Subdocuments via `new_subdoc` may require optional dependencies such as docxcompose, which are not part of the four-tool baseline. Do not assume they are installed. Merged documents still need style/numbering/relationship checks.

`get_undeclared_template_variables()` helps discover top-level requirements; it is not a complete schema or nested-value validator. Validate repeated-row data and required nonempty values before rendering. Compute monetary values with `Decimal` and format them before injecting strings.

Verification cases: missing nested key, `<>&`, Chinese text, empty rows, multiple rows, long wrapping text, images, and page-spanning tables. `docxtpl` does not render or guarantee pagination.

Source: [official documentation](https://docxtpl.readthedocs.io/en/latest/), especially Restrictions, Tags, RichText, Escaping, Replace docx pictures, and Get Defined Variables.
