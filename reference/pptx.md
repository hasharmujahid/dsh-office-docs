# Presentations (.pptx)

Run everything with the skill interpreter: `PY=~/.dsh/office-docs/venv/bin/python`.

## Which approach

| Situation | Approach |
|---|---|
| Create a deck | a `python-pptx` script (patterns below) |
| Build from a template deck's design | open the template with `Presentation("template.pptx")`, then add slides using its layouts |
| Edit existing slides' text | iterate `slide.shapes`, set `text_frame` content in place — never rebuild the deck |
| Read or extract content | `office.py dump deck.pptx` — one block per slide, including speaker notes |
| Visual review of every slide | `office.py render deck.pptx --prefix /tmp/slide --dpi 90` |

A deck is judged visually. **Always render it and look at the pages** before delivering — text overflow,
overlapping shapes, and 8pt body text are invisible in the object model and obvious in the image.

## A correct deck

```python
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

prs = Presentation()                       # 16:9? No: this default is 10 x 7.5 inches (4:3)
prs.slide_width = Inches(13.333)           # set 16:9 explicitly
prs.slide_height = Inches(7.5)

title_slide = prs.slides.add_slide(prs.slide_layouts[0])   # layout 0 = Title Slide
title_slide.shapes.title.text = "Q3 Operations Review"
title_slide.placeholders[1].text = "Prepared for the Board"

content = prs.slides.add_slide(prs.slide_layouts[1])       # layout 1 = Title and Content
content.shapes.title.text = "Highlights"
body = content.placeholders[1].text_frame
body.text = "Revenue up 18%"                               # first paragraph: assignment, not add_paragraph
for line in ["EMEA flat on FX", "APAC strongest at +31%"]:
    para = body.add_paragraph()
    para.text = line
    para.level = 0                                          # 0 = top bullet, 1 = sub-bullet
    para.font.size = Pt(20)

prs.save("deck.pptx")
```

**`prs.slide_layouts` is index-based and template-dependent.** Index 0 is Title Slide and 1 is Title
and Content in the default template, but a corporate template can order them differently. Check first:

```python
for index, layout in enumerate(prs.slide_layouts):
    print(index, layout.name, [ph.placeholder_format.type for ph in layout.placeholders])
```

`prs.slide_width` / `slide_height` are inherited from the template. A deck assembled on a 4:3 canvas
and shown on a 16:9 screen has two black bars, so set both explicitly when starting from default.

## Text fitting — the main source of ugly decks

`python-pptx` does not measure text and does not shrink it to fit. A paragraph longer than its
placeholder overflows past the slide edge. Control it deliberately:

```python
from pptx.util import Pt

frame = content.placeholders[1].text_frame
frame.word_wrap = True                        # wrap instead of running off the shape
frame.auto_size = None                        # never let PowerPoint rescale the shape behind your back

for para in frame.paragraphs:
    for run in para.runs:
        run.font.size = Pt(18)
        run.font.name = "Arial"
```

Keep body text to roughly 6 bullets of ≤ 12 words. When content genuinely does not fit, split it across
two slides or move it into the speaker notes — do not shrink below 16pt.

## Positioning and sizing shapes

```python
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

box = content.shapes.add_textbox(Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.0))
frame = box.text_frame
frame.text = "Source: finance rollup, unaudited"
para = frame.paragraphs[0]
para.alignment = PP_ALIGN.RIGHT
para.runs[0].font.size = Pt(12)

# A colored accent bar: a shape with no text and no outline
from pptx.enum.shapes import MSO_SHAPE
bar = content.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.25), prs.slide_height)
bar.fill.solid()
bar.fill.fore_color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
bar.line.fill.background()
```

Every position and size is an `Inches(...)` / `Pt(...)` / `Emu(...)` value — a bare integer is
interpreted as EMU (914400 per inch), which places shapes effectively at the origin.

## Tables

```python
from pptx.util import Inches, Pt

rows, cols = 3, 3
shape = content.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(2.4))
table = shape.table
table.columns[0].width = Inches(5.0)          # column widths are set on the column
table.rows[0].height = Inches(0.5)

for col, header in enumerate(["Region", "Revenue", "Growth"]):
    cell = table.cell(0, col)
    cell.text = header
    cell.text_frame.paragraphs[0].runs[0].font.bold = True
    cell.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
```

A table with more rows than fit overflows the slide silently. Prefer fewer rows per slide.

## Charts, images, and notes

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

chart_data = CategoryChartData()
chart_data.categories = ["Q1", "Q2", "Q3"]
chart_data.add_series("Revenue", (120, 145, 168))
content.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.5), chart_data
)

# Images: always constrain one dimension, or a photo lands at native pixel size
content.shapes.add_picture("logo.png", Inches(10.5), Inches(6.4), height=Inches(0.6))

# Speaker notes
content.notes_slide.notes_text_frame.text = "Mention the FX headwind, then move to APAC."
```

Charts created this way hold their data in the package, so they render in LibreOffice and PowerPoint
without needing the source workbook.

## Editing an existing deck

`python-pptx` can open an existing `.pptx` and modify text in place, which is the right approach for
adding a slide to someone else's deck:

```python
prs = Presentation("existing.pptx")
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.paragraphs:
                for run in para.runs:
                    if "TBD" in run.text:
                        run.text = run.text.replace("TBD", "Confirmed")
prs.save("updated.pptx")
```

Two cautions. Replacing text across runs has the same fragmentation problem as Word: a phrase split
across two runs will not match. And **`python-pptx` re-serialises unknown parts it does not model**,
so a heavily customised deck (SmartArt, embedded video, custom XML) can lose pieces on save. When a
deck is complex, prefer editing the XML inside `ppt/slides/slideN.xml` and rezipping, keeping every
other part byte-identical.

Deleting a slide is not a first-class operation: remove the `<p:sldId>` element from
`ppt/presentation.xml`'s `<p:sldIdLst>` and drop the slide part, or reorder by rewriting that list.
Duplicate a slide by copying its part and adding a matching relationship and `<p:sldId>` entry — the
package bookkeeping is what makes this fiddly, so verify by rendering.

## Gotchas

| Symptom | Cause |
|---|---|
| Text runs past the slide edge | no autofit in `python-pptx`; shorten the text or split the slide |
| Shape appears in the top-left corner | a bare integer size/position is EMU, not points |
| `KeyError` or `IndexError` on `placeholders[1]` | that layout has a different placeholder set; print the layout before indexing |
| Black bars beside every slide | 4:3 default vs a 16:9 screen; set `slide_width` to 13.333" |
| Bold or font size ignored | setting it on the paragraph instead of the `run`, or the theme overrides it |
| `add_paragraph()` then `.text =` on the frame | assignment replaces the whole frame; use one or the other |

## Verify

```bash
$PY "$SKILL/scripts/office.py" validate deck.pptx
$PY "$SKILL/scripts/office.py" render deck.pptx --prefix /tmp/slide --dpi 90
$PY "$SKILL/scripts/office.py" dump deck.pptx          # text and notes, per slide
```

Read the rendered images. For a deck longer than a few slides, render at `--dpi 80` and read the pages
in batches — a full-size deck at high dpi floods the context for no extra signal.
