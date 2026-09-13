# office-docs

[![test](https://github.com/OWNER/REPO/actions/workflows/test.yml/badge.svg)](https://github.com/OWNER/REPO/actions/workflows/test.yml)

A DSH skill that makes Word, Excel, and PowerPoint deliverables reliable: build them, edit them under
tracked changes, validate them, and *look* at the result before handing it over.

It exists because the failure mode of "let the model write OOXML from memory" is not a crash — it is a
file that opens and is subtly wrong: A4 instead of Letter, a table that collapses elsewhere, a
deletion that renders as unchanged text, an edit that survives a reviewer's "reject all".

## What is in here

```
office-docs/
├── SKILL.md                       # routing, house rules, troubleshooting (loaded by DSH)
├── reference/
│   ├── docx.md                    # python-docx + OOXML guidance, tested recipes
│   ├── xlsx.md                    # openpyxl: formulas, charts, number formats
│   └── pptx.md                    # python-pptx: layouts, fitting, charts
└── scripts/
    ├── bootstrap.py               # create the shared venv (idempotent, --check/--force)
    ├── office.py                  # info · dump · convert · render · new · validate
    ├── md2docx.py                 # the Markdown -> .docx renderer used by `new`
    ├── docx_edit.py               # merge-runs · find-replace --track · comment · accept · reject
    ├── smoke_test.py              # 27 end-to-end checks of the real toolchain
    └── verify_reference_code.py   # executes every complete example in reference/
```

## Install

```bash
./install.sh                      # -> ~/.dsh/skills/office-docs + shared venv
./install.sh --link               # symlink, for editing this copy in place
./install.sh --project ~/work/x   # -> ~/work/x/.dsh/skills/office-docs
```

DSH discovers a skill at `<root>/<name>/SKILL.md`; the sibling `scripts/` and `reference/` directories
travel with it as the skill's resources. Recognised roots are `<project>/.dsh/skills`,
`<project>/.agents/skills`, `$DSH_HOME/skills`, and `$DSH_AGENTS_HOME/skills`.

## Use

```bash
PY=~/.dsh/office-docs/venv/bin/python
SKILL=~/.dsh/skills/office-docs

$PY "$SKILL/scripts/smoke_test.py"                     # is the toolchain healthy?

# Markdown -> Word
$PY "$SKILL/scripts/office.py" new report.md --title "Q3 Review" --toc --page-numbers

# Redline an existing document under a reviewer's name
$PY "$SKILL/scripts/docx_edit.py" find-replace in.docx out.docx \
    --old "ACME" --new "Acme Corporation" --author "A. Reviewer"
$PY "$SKILL/scripts/office.py" validate out.docx --original in.docx --author "A. Reviewer"

# Look at the result before delivering it
$PY "$SKILL/scripts/office.py" render out.docx --prefix /tmp/page --dpi 100
```

## How it stays reliable

- **One environment, pinned majors.** `bootstrap.py` builds `~/.dsh/office-docs/venv` with
  `python-docx`, `openpyxl`, `python-pptx`, `lxml`, and `pillow` at stable major versions. A floating
  install is what breaks a document pipeline six months after it was written.
- **Validation is a command, not a hope.** `office.py validate` checks package integrity, every XML
  part, formula error literals, comment cross-links, revision consistency, whether an edit bypassed
  tracked changes, and whether LibreOffice can open the file at all.
- **Rendering closes the loop.** `office.py render` converts to PDF via LibreOffice and rasterizes with
  poppler so a model can actually see the pages it produced.
- **The traps are documented where they bite.** Every pattern in `reference/` is executed by
  `verify_reference_code.py`, so the docs cannot quietly rot.

## Requirements

Python 3.9+ (3.14 used here) · network access for the one-time `pip install` · LibreOffice
(`libreoffice-writer`, `-calc`, `-impress`) and `poppler-utils` for conversion, PDF export, and
rendering. The skill degrades gracefully: without LibreOffice, `new`, `dump`, `find-replace`,
`comment`, and `accept`/`reject` still work.

```bash
# Debian/Ubuntu
sudo apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress poppler-utils
# macOS
brew install --cask libreoffice && brew install poppler
```

## Repository layout

```
SKILL.md                       the file DSH loads
reference/{docx,xlsx,pptx}.md  format guidance; executable examples
scripts/
  bootstrap.py                 create the shared environment
  office.py                    info · dump · convert · render · new · validate
  md2docx.py                   Markdown -> .docx renderer
  docx_edit.py                 tracked changes, comments, accept/reject
  smoke_test.py                27 end-to-end checks
  verify_reference_code.py     runs every complete example in reference/
install.sh                     install into a DSH skills root
CONTRIBUTING.md                invariants and how to change this safely
```

## Development

CI runs the smoke test and the reference-code verification on Python 3.9 and 3.12 with LibreOffice and
poppler installed, so a change that only works on one machine fails there. See
[CONTRIBUTING.md](CONTRIBUTING.md) before editing the XML handling — the invariants listed there are
each the result of a silently wrong document.

## License

MIT — see [LICENSE](LICENSE).

The design was informed by [anthropics/skills](https://github.com/anthropics/skills) (proprietary, not
copied): its `docx`/`xlsx`/`pptx` skills established the shape of the workflow. Every line here is an
independent implementation against `python-docx`, `openpyxl`, `python-pptx`, and LibreOffice, because
the reference stack's dependencies (`pandoc`, npm `docx`, `markitdown`) are not present in this
environment.

