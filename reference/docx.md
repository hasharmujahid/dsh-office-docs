# Word documents (.docx)

Run everything with the skill interpreter: `PY=~/.dsh/office-docs/venv/bin/python`.

## Which approach

| Situation | Approach |
|---|---|
| New document from prose | Write Markdown, then `office.py new` (see SKILL.md) |
| New document needing precise layout | A `python-docx` script, as below |
| Any change to an existing document | `docx_edit.py` (find-replace / comment), never a rebuild |
| Read or extract content | `office.py dump file.docx` |
| Content in a template's branding | `office.py new --template corporate.docx` |

The Markdown path exists because a model writing `python-docx` calls from memory reliably produces the
same handful of defects: A4 instead of Letter, literal `•` characters instead of real list numbering,
tables that collapse in Google Docs, and headings that never appear in a table of contents. If you do
write a script, copy the patterns below rather than recalling the API.

## A correct new document

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

doc = Document()                      # US Letter, 1" margins, Calibri 11
style = doc.styles["Normal"]          # set the body font once, not per run
style.font.name = "Arial"
style.font.size = Pt(11)

title = doc.add_heading("Quarterly Review", level=0)   # level=0 is the Title style
title.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.add_heading("Findings", level=1)  # real Heading styles: these reach the TOC
para = doc.add_paragraph("Revenue grew ")
para.add_run("18%").bold = True       # formatting lives on runs, not paragraphs
para.add_run(" quarter over quarter.")

doc.save("out.docx")
```

**Page setup is US Letter by default in `python-docx`.** If you need A4 explicitly, set
`section.page_width = Mm(210)`, `section.page_height = Mm(297)`. For landscape, swap the width and
height *and* set `section.orientation = WD_ORIENT.LANDSCAPE` — setting orientation alone leaves the
text rotated inside a portrait page.

## Tables

Column widths are only respected when the table is fixed-layout and every cell carries a width. Set
both, or Word and Google Docs disagree about the result:

```python
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

table = doc.add_table(rows=1, cols=3)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False                       # required, or Word re-flows the columns

for cell, text in zip(table.rows[0].cells, ["Region", "Revenue", "Growth"]):
    cell.text = ""
    run = cell.paragraphs[0].add_run(text)
    run.bold = True
    shd = OxmlElement("w:shd")              # shading: w:fill, never a solid pattern
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "E8EDF4")
    cell._tc.get_or_add_tcPr().append(shd)

for cell, width in zip(table.rows[0].cells, [Inches(2.5), Inches(2.0), Inches(2.0)]):
    cell.width = width                      # width on the CELL, not just the column
```

`table.style = "Table Grid"` is one of the styles guaranteed to exist. A custom style name that is
not in the file raises `KeyError`.

## Lists

Use the built-in list styles. Inserting a literal `•` produces text that will not renumber and looks
wrong in any other renderer:

```python
for item in ["North America", "EMEA", "APAC"]:
    doc.add_paragraph(item, style="List Bullet")
for item in ["Supplier concentration", "Renewal cycle"]:
    doc.add_paragraph(item, style="List Number")   # numbers are computed by Word
```

Nesting uses `List Bullet 2` / `List Number 2`, and so on.

## Images

```python
doc.add_picture("chart.png", width=Inches(6.0))     # width only: height scales
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
```

Give a width (or height) — an image added without one lands at its native pixel size and usually
overflows the margin. `docx_edit.py` does not touch images; for replacing an image inside an existing
document, extract `word/media/` and swap the file, keeping the same part name.

## Headers, footers, page numbers

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_page_numbers(doc):
    footer = doc.sections[0].footer
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    for element, text in (("begin", None), ("instrText", "PAGE"), ("separate", None), ("end", None)):
        if element == "instrText":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), element)
        run._r.append(el)
```

`PAGE` is the current page, `NUMPAGES` the total. Word shows the field result only after opening the
document; a placeholder until then is expected. `office.py new --page-numbers` does this for you.

The same field technique builds a table of contents (`TOC \o "1-3" \h \z \u`). **A TOC needs heading
styles**: a heading formatted manually with bold text and a larger font is invisible to it.

## Equations

Word cannot be handed LaTeX. `office.py new` translates simple expressions (`$E = mc^2$` →
*E = mc²*) into editable italic runs. For real equations, insert them from Word's equation editor, or
extract the OMML from an existing document and re-insert the `<m:oMath>` element — do not paste LaTeX
into a run and call it an equation.

## Editing an existing document

Direct XML edits are the correct approach for an existing file — `python-docx` cannot round-trip one
without dropping parts it does not model. Two rules make it safe:

1. **Never pretty-print or reindent `document.xml`.** Whitespace inside `<w:t xml:space="preserve">`
   is content.
2. **Merge runs before searching.** Word splits a sentence across many `<w:r>` elements, so the phrase
   you can see is usually not a contiguous string. `docx_edit.py` merges automatically; by hand, use
   `docx_edit.py merge-runs`.

```bash
$PY "$SKILL/scripts/docx_edit.py" merge-runs in.docx -o merged.docx
$PY "$SKILL/scripts/docx_edit.py" find-replace in.docx out.docx --old "old" --new "new" --author "Me"
$PY "$SKILL/scripts/docx_edit.py" accept in.docx clean.docx     # accept all revisions
$PY "$SKILL/scripts/docx_edit.py" reject in.docx original.docx  # reject all revisions
```

A legacy `.doc` must be converted first: `office.py convert file.doc --to docx`.

## XML details that bite

- Deleted text inside a revision uses `<w:delText>`, **not** `<w:t>`. Get this wrong and the deletion
  renders as normal text.
- In a `<w:rPr>`, `<w:del/>` must precede the other children — the element order is schema-enforced.
- A deleted paragraph mark is `<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>`, which means "join this
  paragraph to the next". Deleting a paragraph outright means that mark *plus* a `<w:del>` around
  every run.
- Comments need six cross-linked parts. Use `docx_edit.py comment`; hand-written comment XML produces
  a file Word refuses to open or a comment with no visible anchor.
- Extraction regexes must distinguish `<w:t>` from `<w:delText>` and `<w:instrText>`. The pattern
  `<w:t[^>]*>` matches all three and silently mixes deleted text into your output. The correct shape is
  `<w:t(?:\s[^>]*)?>`.

## Verify

```bash
$PY "$SKILL/scripts/office.py" validate out.docx
$PY "$SKILL/scripts/office.py" render out.docx --prefix /tmp/page --dpi 100
```

Then read the rendered pages. For an edited document, always validate against the original
(`--original in.docx --author "Name"`); that is the only check that catches an edit which bypassed
tracked changes.
