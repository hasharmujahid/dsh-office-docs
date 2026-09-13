# Spreadsheets (.xlsx)

Run everything with the skill interpreter: `PY=~/.dsh/office-docs/venv/bin/python`.

## Which approach

| Situation | Approach |
|---|---|
| Create or format a workbook | an `openpyxl` script (patterns below) |
| Bulk data in or out | `pandas.read_excel` / `to_excel` — then reformat with `openpyxl` if it must look designed |
| Read an existing workbook | `office.py dump file.xlsx` — prints every sheet with formulas *and* cached values |
| Read values for computation | `openpyxl.load_workbook(path, data_only=True)` |
| Convert to CSV / PDF | `office.py convert file.xlsx --to csv` |

**The single most common failure is writing data with no formulas.** A spreadsheet whose totals are
baked-in numbers cannot recalculate when a reader changes an input, and the reader will change an
input. Always write `"=SUM(B2:B9)"`, never the Python-computed total.

## A well-formed workbook

```python
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Revenue"                     # sheet names are part of the deliverable

headers = ["Region", "Q1", "Q2", "Total"]
for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F3A5F")
    cell.alignment = Alignment(horizontal="center", vertical="center")

rows = [("North America", 120, 145), ("EMEA", 90, 88), ("APAC", 70, 95)]
for idx, (region, q1, q2) in enumerate(rows, start=2):
    ws.cell(row=idx, column=1, value=region).font = Font(name="Arial")
    ws.cell(row=idx, column=2, value=q1).font = Font(name="Arial")
    ws.cell(row=idx, column=3, value=q2).font = Font(name="Arial")
    ws.cell(row=idx, column=4, value=f"=B{idx}+C{idx}").font = Font(name="Arial")

last = len(rows) + 1
total = ws.cell(row=last + 1, column=1, value="Total")   # the label, then the formula
total.font = Font(name="Arial", bold=True)
ws.cell(row=last + 1, column=4, value=f"=SUM(D2:D{last})").font = Font(name="Arial", bold=True)

# Number formats: money with a currency symbol, percentages as integers
for row in range(2, last + 1):
    ws.cell(row=row, column=2).number_format = '"$"#,##0'
ws.cell(row=last + 1, column=4).number_format = '"$"#,##0'

# Layout: widths, frozen header, autofilter
ws.column_dimensions["A"].width = 22
for col in "BCD":
    ws.column_dimensions[col].width = 11
ws.freeze_panes = "A2"                   # freezes row 1 while scrolling
ws.auto_filter.ref = f"A1:D{last}"

thin = Side(style="thin", color="D0D7E5")
for row in ws[f"A1:D{last + 1}"]:
    for cell in row:
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb.save("revenue.xlsx")
```

`Font(name="Arial")` must be set on every cell you write — a workbook with Calibri defaults next to a
few Arial cells looks accidental.

## Reading: the `data_only` trap

A formula cell holds either the formula string or the last value Excel cached, never both:

```python
import openpyxl

formulas = openpyxl.load_workbook("in.xlsx", data_only=False)   # "=SUM(B2:B9)"
values = openpyxl.load_workbook("in.xlsx", data_only=True)      # 1370, or None

print(formulas["Revenue"]["D2"].value, values["Revenue"]["D2"].value)
```

**A file created by a library has no cached values at all**, so `data_only=True` returns `None` for
every formula until Excel or LibreOffice recalculates and saves it. If you need the computed number,
either calculate it in Python or convert through LibreOffice first:

```bash
$PY "$SKILL/scripts/office.py" convert in.xlsx --to xlsx --outdir /tmp/recalc
```

`None` from `data_only=True` is therefore not evidence that a formula is broken. `office.py validate`
loads the file a second time without `data_only` for exactly this reason.

## Multiple sheets, references, named ranges

```python
summary = wb.create_sheet("Summary")                 # appended after the active sheet
summary["B2"] = "=SUM(Revenue!D2:D4)"                # sheet names with spaces need quotes: 'Q1 Data'!A1
wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName("TaxRate", attr_text="Settings!$B$1"))
summary["B3"] = "=B2*TaxRate"                        # referencing a defined name
sheet = wb["Revenue"]                                # sheet access by name
del wb["Scratch"]                                    # remove a sheet you no longer need
```

Sheet order in `wb.sheetnames` is display order. Renaming a sheet after writing cross-sheet formulas
breaks them — rename first, then write formulas.

## Charts

```python
from openpyxl.chart import BarChart, Reference

chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "Revenue by region"
chart.y_axis.title = "USD"
chart.x_axis.title = "Region"

data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last)   # includes the header row
categories = Reference(ws, min_col=1, min_row=2, max_row=last)
chart.add_data(data, titles_from_data=True)      # titles_from_data picks up the header row
chart.set_categories(categories)
chart.height, chart.width = 8, 16                # centimetres
ws.add_chart(chart, "F2")                        # anchor cell for the top-left corner
```

Anchor charts to an empty region — an anchored chart floats over whatever occupies those cells.
Charts survive a `load_workbook` / `save` round-trip, but **a chart whose data was written without
cached values renders empty** until the file is opened and recalculated once.

## Conditional formatting, data validation, comments

```python
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

ws.conditional_formatting.add(
    f"D2:D{last}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FDECEA"))
)
ws.conditional_formatting.add(f"B2:C{last}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="8FD19E"))

validation = DataValidation(type="list", formula1='"Draft,Review,Final"', allow_blank=True)
ws.add_data_validation(validation)
validation.add("E2:E50")

ws["A1"].comment = Comment("Source: finance rollup", "Analyst", width=200, height=60)
```

## Gotchas

| Symptom | Cause |
|---|---|
| `KeyError` on `wb["Sheet1"]` | the default sheet is renamed when you set `ws.title`; track the object, not the name |
| Totals are wrong after editing | a hardcoded total: change the writer to emit `=SUM(...)` |
| `#REF!` after deleting rows | formulas were written as strings with wrong offsets; `office.py validate` reports these |
| Column shows `#####` | column too narrow or a negative value in a date format; widen it or fix `number_format` |
| Dates stored as text | assign `datetime.date`/`datetime` objects, then set `number_format = "yyyy-mm-dd"` |
| New sheet appears before others | `wb.create_sheet("Name", 0)` inserts at index 0; default appends |
| Formula error values appear | `office.py validate` reports `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#NUM!`, `#N/A` |

For very large sheets, `load_workbook(..., read_only=True)` streams rows and is dramatically faster —
but the returned cells are immutable, so it is a read path only. `write_only=True` streams a write.
Neither supports formatting the way the normal mode does.

## Verify

```bash
$PY "$SKILL/scripts/office.py" validate out.xlsx      # includes formula-error detection
$PY "$SKILL/scripts/office.py" dump out.xlsx          # formulas and cached values, per sheet
$PY "$SKILL/scripts/office.py" convert out.xlsx --to pdf   # then render and read the pages
```

For a spreadsheet a human will open, render it and look: a mis-sized column or an unformatted header
row is obvious in the image and invisible in the data.
