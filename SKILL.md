---
name: office-docs
description: Use when a deliverable is a Word, Excel, or PowerPoint file (.docx/.dotx, .xlsx/.xlsm, .pptx/.potx) — creating reports, memos, letters, or templates; editing or redlining an existing document with tracked changes or comments; building or fixing spreadsheets with formulas, formatting, and charts; making or editing slide decks; extracting content from an Office file; or converting between Office formats and PDF. Triggers include "Word doc", "spreadsheet", "deck", "slides", "redline", "tracked changes", "fill in this template", and any path ending in those extensions. Do not use for plain Markdown or code files, or when the deliverable is a PDF that is not built from an Office document.
whenToUse: The user asks for a Word document, spreadsheet, or slide deck as the deliverable, or asks to read, edit, redline, or convert an existing .docx/.xlsx/.pptx.
---

# Office documents (docx · xlsx · pptx)

Everything here runs through one shared Python environment. **Use its interpreter, not `python3`** — the
system Python does not have the document libraries:

```bash
PY=~/.dsh/office-docs/venv/bin/python          # the interpreter for every command below
SKILL=${DSH_HOME:-$HOME/.dsh}/skills/office-docs   # this skill's directory
```

If `$PY` does not exist yet, create it once (needs network):

```bash
python3 "$SKILL/scripts/bootstrap.py"
```

Before trusting the toolchain — after an install, a Python upgrade, or any odd result — run the self-test:

```bash
$PY "$SKILL/scripts/smoke_test.py"      # 27 checks: build, edit, validate, convert, render
```

## Pick the operation, then run the script

| The task | Use this | Never |
|---|---|---|
| Create a Word document from prose | write Markdown → `office.py new` | hand-write OOXML |
| Edit an existing `.docx` | `docx_edit.py find-replace` / `comment` | rewrite the whole file |
| Redline under a reviewer's name | `docx_edit.py … --author "Name"` | untracked edits |
| Create/edit a spreadsheet | a short `openpyxl` script | write `.xlsx` as CSV or text |
| Create/edit a deck | a short `python-pptx` script | edit `slideN.xml` by hand |
| Read any Office file | `office.py dump` | unzip and read XML by eye |
| Convert to PDF/other | `office.py convert --to pdf` | a different converter tool |
| Deliver with confidence | `office.py render` then **look at the pages** | assume it looks right |

## The loop that makes output reliable

1. **Build** the file with a script or `office.py new`.
2. **Validate** it: `office.py validate out.docx` — checks package integrity, XML well-formedness,
   formula errors, comment cross-links, revision consistency, and whether LibreOffice can open it.
3. **Render** it: `office.py render out.docx --prefix /tmp/page` → then read the PNGs with the image
   tool. This is the step that catches a table that overflows the margin, a heading that never got a
   heading style, or text sitting on top of an image. Do not skip it for anything a human will open.
4. **Fix and repeat**, then hand over the file with `present`.

## House rules for every deliverable

- **Real content, not placeholders.** If the user's spec names tabs, headers, or wording, follow it literally.
- **Formulas, never hardcoded results.** `sheet["D2"] = "=B2+C2"` — the sheet must recalculate when inputs change.
- **Zero formula errors before delivery.** `office.py validate` fails on `#REF!`, `#DIV/0!`, and friends.
- **One professional font throughout** (Arial or Times New Roman) unless the user asks otherwise.
- **Rendering for visual review needs LibreOffice + poppler.** Both are present here; see Troubleshooting.
- **Keep scratch files out of the user's directory.** Render pages to a temp dir, not next to the deliverable.

## Tracked changes and comments

An edit written straight into the XML is invisible to a reviewer's "reject all" — from the accepted
view it looks exactly like a tracked edit. So redlining is a workflow, not a flag:

```bash
# Redline an existing document: every change carries the author name.
$PY "$SKILL/scripts/docx_edit.py" find-replace in.docx out.docx \
    --old "ACME Corp" --new "Acme Corporation" --author "A. Reviewer"

# Anchor a comment to specific existing text (without an anchor it exists but is invisible).
$PY "$SKILL/scripts/docx_edit.py" comment in.docx -o out.docx \
    --text "Cap is too low; propose 25%." --anchor "Supplier concentration" --author "A. Reviewer"

# Prove it: the validator compares against the original and fails on any untracked difference.
$PY "$SKILL/scripts/office.py" validate out.docx --original in.docx --author "A. Reviewer"
```

`w:del` inside `<w:rPr>` with no `<w:delText>` in the run is a real defect: the deletion renders as
unchanged text. The validator reports it.

## Format-specific guidance

Read the reference file for the format you are working in — the API traps live there:

- **Word** → `reference/docx.md` — building documents, styles that reach the TOC, tables and images,
  templates, footers, equations, and the run-splitting problem.
- **Excel** → `reference/xlsx.md` — formulas, number formats, charts, column widths, frozen panes,
  and the `data_only` trap that makes values vanish.
- **PowerPoint** → `reference/pptx.md` — layouts and placeholders, text fitting, tables, charts,
  slide geometry, and speaker notes.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `ModuleNotFoundError: docx` | wrong interpreter — use `$PY`, or run `scripts/bootstrap.py` |
| LibreOffice `DeploymentException` or `Application Error` | a read-only or shared user profile; `office.py` avoids it with `-env:UserInstallation`. Do not call `soffice` directly |
| `soffice` not found | `apt-get install libreoffice-writer libreoffice-calc libreoffice-impress` |
| `pdftoppm` not found | `apt-get install poppler-utils`; only rendering needs it |
| Render is blank or missing pages | convert to PDF first and check `pdfinfo`; some decks export empty pages |
| Word says the file is unreadable | a comment or revision part is unlinked; compare with `office.py validate` |
| Find-and-replace reports no match | the text spans runs, punctuation differs (curly quotes), or it is inside a text box; `merge-runs` first, then search a shorter phrase |
| TOC field shows the placeholder text | fields update when opened in Word; that is expected, not a defect |

## Installing this skill elsewhere

The skill directory is self-contained: copy `office-docs/` into any DSH skills root
(`~/.dsh/skills/`, `~/.agents/skills/`, or `<project>/.dsh/skills/`) and run `bootstrap.py` once on the
new machine. The environment default is `~/.dsh/office-docs/venv`; override it with `OFFICE_DOCS_VENV`.
