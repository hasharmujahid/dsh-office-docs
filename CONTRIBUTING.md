# Contributing

## Setup

```bash
python3 scripts/bootstrap.py                     # creates ~/.dsh/office-docs/venv
~/.dsh/office-docs/venv/bin/python scripts/smoke_test.py
```

`bootstrap.py` is idempotent — run it any time and it verifies instead of reinstalling. `--check`
verifies without network access; `--force` rebuilds from scratch.

## Before you open a pull request

```bash
~/.dsh/office-docs/venv/bin/python scripts/smoke_test.py            # 27 checks, must be all PASS
~/.dsh/office-docs/venv/bin/python scripts/verify_reference_code.py # every doc example must run
```

CI runs both, on Python 3.9 and 3.12, with LibreOffice and poppler installed. A change that only works
on the author's machine fails there, which is the point.

## Invariants

These are load-bearing. Each one exists because breaking it produced a silently wrong document.

1. **`<w:t(?:\s[^>]*)?>` — never `<w:t[^>]*>`.** The loose form also matches `<w:delText>` and
   `<w:instrText>`, so deleted text and field codes leak into extracted output as if they were body
   text. This is the highest-consequence one-line change possible in this repository.

2. **`<w:ins\b` and `<w:del(?:\s[^>]*)?\s*/?>` for counting revisions.** A bare `<w:ins` also matches
   `<w:instrText>`; a bare `<w:del` also matches `<w:delText>`.

3. **Never match deletions with `<w:del\b.*?</w:del>`.** Word writes a self-closing `<w:del/>` inside
   `w:pPr` for a deleted paragraph mark, which makes that pattern run on to the *next* `</w:del>` and
   swallow everything between them, insertions included.

4. **`--author` is mandatory for tracked edits.** An untracked edit is indistinguishable from a tracked
   one in the accepted view, so `office.py validate --original` is the only thing that catches it. Do
   not weaken that check to make a failing document pass.

5. **Comments need their anchor.** A comment part with no `commentRangeStart` in `document.xml` exists
   but is invisible in Word. `docx_edit.py comment` refuses to write an anchor it cannot find; keep
   that behaviour.

6. **No shared internal module between `office.py`, `md2docx.py`, and `docx_edit.py`.** Each is meant
   to be copyable into a project on its own, which is worth the small duplication.

## Adding a reference recipe

Add the code to `reference/<format>.md` inside a ` ```python ` fence, then run
`verify_reference_code.py`. A block that both imports something and calls `.save(...)` is executed
automatically — so a recipe that does not work fails CI rather than misleading someone later.
Fragments that depend on a document from a previous block are skipped and should stay small.

## Style

- Standard library plus the five pinned dependencies. New dependencies need a reason in the PR.
- Every script is runnable directly and has a module docstring explaining the failure it prevents,
  not just what it does.
- Error messages name the fix. `ERROR: no such file: x.docx` is worse than telling the caller which
  command creates it.
- Prefer a deterministic script over a paragraph of instructions to the model.

## Releasing

There is no release artifact; the skill is the source tree. Tag a version when the Docker-free
behaviour changes in a way users would notice:

```bash
git tag -a v0.2.0 -m "office-docs 0.2.0" && git push origin v0.2.0
```
