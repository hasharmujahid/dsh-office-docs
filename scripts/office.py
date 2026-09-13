#!/usr/bin/env python3
"""office.py - deterministic document operations for the office-docs skill.

One entry point for the operations that are easy to get subtly wrong when an
agent improvises them: format conversion, page rendering for visual review,
content extraction, structure inspection, and package validation.

Run it with the skill's own interpreter so the document libraries are present:

    ~/.dsh/office-docs/venv/bin/python scripts/office.py render report.docx

`basic` subcommands (info, dump, convert, render) work on the standard library
plus LibreOffice/pdftoppm. `new` and `validate` need the Python libraries; if
they are missing, this script tells you the exact command that fixes it.

Exit codes: 0 success · 1 usage/verification failure · 2 missing dependency
or missing external tool.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

EXIT_OK = 0
EXIT_FAIL = 1
EXIT_MISSING = 2

DEFAULT_ROOT = Path(os.environ.get("OFFICE_DOCS_HOME", Path.home() / ".dsh" / "office-docs"))
DEFAULT_VENV = Path(os.environ.get("OFFICE_DOCS_VENV", DEFAULT_ROOT / "venv"))

# Formats LibreOffice can reliably write from an office source document.
CONVERT_TARGETS = {
    "pdf": "pdf",
    "docx": "docx",
    "xlsx": "xlsx",
    "pptx": "pptx",
    "txt": "txt",
    "csv": "csv",
    "rtf": "rtf",
    "odt": "odt",
    "ods": "ods",
    "odp": "odp",
    "html": "html",
    "png": "png",
}


def venv_python(venv: Path) -> Path:
    return venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def require_libs(what: str) -> None:
    """Fail with an actionable message when the shared venv is absent."""
    probe = subprocess.run(
        [sys.executable, "-c", "import docx, openpyxl, pptx"],
        capture_output=True,
        text=True,
    )
    if probe.returncode == 0:
        return
    boot = Path(__file__).resolve().parent / "bootstrap.py"
    print(
        f"ERROR: {what} needs the document libraries, which are not importable by\n"
        f"       {sys.executable}\n\n"
        f"Fix it once with:\n  python3 {boot}\n\n"
        f"then re-run this command with {venv_python(DEFAULT_VENV)}",
        file=sys.stderr,
    )
    raise SystemExit(EXIT_MISSING)


def require_tool(name: str, hint: str) -> str:
    found = shutil.which(name)
    if found is None and name == "soffice":
        found = shutil.which("libreoffice")
    if found is None:
        print(f"ERROR: `{name}` not found on PATH. {hint}", file=sys.stderr)
        raise SystemExit(EXIT_MISSING)
    return found


def soffice(*args: str, timeout: int = 300) -> subprocess.CompletedProcess:
    """Run LibreOffice headless with an isolated, writable profile.

    A shared or read-only user profile is the usual cause of the opaque
    `DeploymentException` / `Application Error` abort, so every call gets its
    own profile directory under a writable HOME.
    """
    binary = require_tool(
        "soffice",
        "Install LibreOffice, or use the format-specific command instead.",
    )
    profile = Path(tempfile.mkdtemp(prefix="office-docs-lo-"))
    home = profile / "home"
    home.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env["HOME"] = str(home)
    env.pop("XDG_RUNTIME_DIR", None)
    cmd = [
        binary,
        f"-env:UserInstallation=file://{profile / 'profile'}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        *args,
    ]
    try:
        return subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=timeout)
    except subprocess.TimeoutExpired:
        print(f"ERROR: LibreOffice timed out after {timeout}s", file=sys.stderr)
        raise SystemExit(EXIT_FAIL)
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def noisy(text: str) -> str:
    """Drop the dconf/Java noise LibreOffice emits in containers."""
    keep = [
        line
        for line in (text or "").splitlines()
        if line.strip()
        and "dconf-CRITICAL" not in line
        and "javaldx" not in line
        and "WARNING: no suitable java" not in line
    ]
    return "\n".join(keep)


def err(msg: str) -> int:
    print(f"ERROR: {msg}", file=sys.stderr)
    return EXIT_FAIL


# ── info ────────────────────────────────────────────────────────────────────


def cmd_info(args: argparse.Namespace) -> int:
    path = Path(args.file)
    if not path.exists():
        return err(f"no such file: {path}")
    print(f"file       : {path}")
    print(f"size       : {path.stat().st_size:,} bytes")
    if not zipfile.is_zipfile(path):
        print("kind       : not an OOXML package (legacy binary or plain text)")
        if args.json:
            print(json.dumps({"path": str(path), "kind": "legacy"}, indent=2))
        return EXIT_OK

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        bad = zf.testzip()
        print(f"kind       : OOXML package ({len(names)} parts)")
        if bad is not None:
            print(f"CORRUPT    : first bad member: {bad}")
        parts_of_interest = [
            n
            for n in names
            if n in {
                "word/document.xml",
                "xl/workbook.xml",
                "ppt/presentation.xml",
                "docProps/core.xml",
                "docProps/app.xml",
            }
            or n.startswith(("word/comments", "word/people", "xl/comments"))
        ]
        for name in sorted(parts_of_interest):
            print(f"  part     : {name} ({zf.getinfo(name).file_size:,} bytes)")
        media = [n for n in names if n.startswith(("word/media/", "ppt/media/", "xl/media/"))]
        if media:
            print(f"media      : {len(media)} embedded file(s)")
        core = zf.read("docProps/core.xml").decode("utf-8", "replace") if "docProps/core.xml" in names else ""
        if core:
            import re

            title = re.search(r"<dc:title>(.*?)</dc:title>", core)
            author = re.search(r"<dc:creator>(.*?)</dc:creator>", core)
            if title and title.group(1):
                print(f"title      : {title.group(1)}")
            if author and author.group(1):
                print(f"author     : {author.group(1)}")

        if args.json:
            print(json.dumps({"path": str(path), "kind": "ooxml", "parts": names}, indent=2))

    if path.suffix.lower() == ".docx":
        try:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            import md2docx  # noqa: WPS433 - optional, only for docx stats

            stats = md2docx.docx_stats(path)
            print(
                f"docx       : {stats['paragraphs']} paragraphs, {stats['tables']} tables, "
                f"{stats['images']} images, {stats['words']} words, "
                f"{stats['tracked_insertions']} tracked insertions, "
                f"{stats['tracked_deletions']} tracked deletions, "
                f"{stats['comments']} comments"
            )
            for h in stats["headings"][:20]:
                print(f"  heading  : {h}")
        except Exception as exc:  # noqa: BLE001 - info must never hard-fail
            print(f"docx       : stats unavailable ({exc})")
    return EXIT_OK


# ── dump / extract ──────────────────────────────────────────────────────────


def cmd_dump(args: argparse.Namespace) -> int:
    """Text extraction that keeps structure, without LibreOffice."""
    require_libs("dump")
    path = Path(args.file)
    if not path.exists():
        return err(f"no such file: {path}")
    suffix = path.suffix.lower()
    out: list[str] = []

    if suffix in {".docx", ".dotx"}:
        import docx
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        doc = docx.Document(str(path))
        body = doc.element.body
        for child in body.iterchildren():
            tag = child.tag.split("}")[-1]
            if tag == "p":
                para = Paragraph(child, doc)
                text = para.text
                style = para.style.name if para.style is not None else ""
                if not text.strip():
                    continue
                if style.startswith("Heading"):
                    level = "".join(c for c in style if c.isdigit()) or "1"
                    out.append(f"{'#' * int(level)} {text}")
                elif style in {"Title", "Subtitle"}:
                    out.append(f"**{text}**")
                elif style.startswith("List"):
                    out.append(f"- {text}")
                else:
                    out.append(text)
            elif tag == "tbl":
                table = Table(child, doc)
                rows = []
                for row in table.rows:
                    rows.append([cell.text.replace("\n", " ").strip() for cell in row.cells])
                out.append(md_table(rows))

    elif suffix in {".xlsx", ".xlsm"}:
        import openpyxl

        wb_formulas = openpyxl.load_workbook(str(path), data_only=False)
        wb_values = openpyxl.load_workbook(str(path), data_only=True)
        for name in wb_formulas.sheetnames:
            ws_f, ws_v = wb_formulas[name], wb_values[name]
            out.append(f"## Sheet: {name}  (dims {ws_f.dimensions}, {ws_f.max_row} rows x {ws_f.max_column} cols)")
            limit_r = args.max_rows if args.max_rows else ws_f.max_row
            limit_c = args.max_cols if args.max_cols else ws_f.max_column
            for r in range(1, min(ws_f.max_row, limit_r) + 1):
                cells = []
                for c in range(1, min(ws_f.max_column, limit_c) + 1):
                    f = ws_f.cell(row=r, column=c).value
                    v = ws_v.cell(row=r, column=c).value
                    if f is None and v is None:
                        cells.append("")
                    elif isinstance(f, str) and f.startswith("="):
                        cells.append(f"{f} => {v!r}")
                    else:
                        cells.append(str(f if f is not None else v))
                while cells and cells[-1] == "":
                    cells.pop()
                if cells:
                    out.append(f"{r}: " + " | ".join(cells))
            if ws_f.max_row > limit_r or ws_f.max_column > limit_c:
                out.append(f"... truncated at {limit_r} rows x {limit_c} cols (raise --max-rows/--max-cols)")

    elif suffix in {".pptx", ".potx"}:
        from pptx import Presentation

        prs = Presentation(str(path))
        out.append(f"# {path.name}: {len(prs.slides)} slides, {prs.slide_width} x {prs.slide_height} EMU")
        for i, slide in enumerate(prs.slides, 1):
            out.append(f"\n<!-- Slide {i} (layout: {slide.slide_layout.name}) -->")
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    for para in shape.text_frame.paragraphs:
                        text = "".join(run.text for run in para.runs)
                        if text.strip():
                            out.append(f"[{shape.shape_type}] {text}")
                elif shape.has_table:
                    rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
                    out.append(md_table(rows))
                elif shape.shape_type == 13:  # PICTURE
                    out.append(f"[picture: {shape.name}]")
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
                out.append(f"NOTES: {slide.notes_slide.notes_text_frame.text.strip()}")
    else:
        return err(f"dump does not handle {suffix}; use `convert --to txt` for {suffix}")

    text = "\n".join(out)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"wrote {args.out} ({len(text):,} chars, {text.count(chr(10)) + 1:,} lines)")
    else:
        print(text)
    return EXIT_OK


def md_table(rows: list[list[str]]) -> str:
    if not rows:
        return "(empty table)"
    width = max(len(r) for r in rows)
    norm = [r + [""] * (width - len(r)) for r in rows]
    header = "| " + " | ".join(norm[0]) + " |"
    sep = "| " + " | ".join("---" for _ in range(width)) + " |"
    body = ["| " + " | ".join(r) + " |" for r in norm[1:]]
    return "\n".join([header, sep, *body])


# ── convert ─────────────────────────────────────────────────────────────────


def cmd_convert(args: argparse.Namespace) -> int:
    src = Path(args.file)
    if not src.exists():
        return err(f"no such file: {src}")
    target = args.to.lower().lstrip(".")
    if target not in CONVERT_TARGETS:
        return err(f"--to must be one of: {', '.join(sorted(CONVERT_TARGETS))}")
    outdir = Path(args.outdir) if args.outdir else src.parent
    outdir.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower().lstrip(".") == target:
        return err(f"source is already .{target}")

    proc = soffice("--convert-to", CONVERT_TARGETS[target], "--outdir", str(outdir), str(src))
    produced = outdir / (src.stem + "." + target)
    if proc.returncode != 0 or not produced.exists():
        print(noisy(proc.stdout), file=sys.stdout)
        print(noisy(proc.stderr), file=sys.stderr)
        return err(f"conversion to .{target} failed")
    print(f"OK  {src}  ->  {produced}  ({produced.stat().st_size:,} bytes)")
    return EXIT_OK


# ── render ──────────────────────────────────────────────────────────────────


def cmd_render(args: argparse.Namespace) -> int:
    """Office file -> PDF -> PNG pages, so the model can actually look at it."""
    src = Path(args.file)
    if not src.exists():
        return err(f"no such file: {src}")
    outdir = Path(args.outdir) if args.outdir else Path(args.prefix).parent
    prefix = Path(args.prefix).name
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = src.with_suffix(".pdf")
    cleanup_pdf = False
    if src.suffix.lower() != ".pdf":
        proc = soffice("--convert-to", "pdf", "--outdir", str(outdir), str(src))
        pdf = outdir / (src.stem + ".pdf")
        if proc.returncode != 0 or not pdf.exists():
            print(noisy(proc.stdout), file=sys.stdout)
            return err(f"could not produce a PDF from {src}")
        cleanup_pdf = not args.keep_pdf

    pdftoppm = require_tool("pdftoppm", "Install poppler-utils for page rendering.")
    fmt = "png" if args.format == "png" else "jpeg"
    cmd = [pdftoppm, f"-{fmt}", "-r", str(args.dpi)]
    if args.pages:
        first, _, last = args.pages.partition("-")
        cmd += ["-f", first, "-l", last or first]
    cmd += [str(pdf), str(outdir / prefix)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(noisy(proc.stderr), file=sys.stderr)
        return err("pdftoppm failed")

    ext = "png" if fmt == "png" else "jpg"
    pages = sorted(outdir.glob(f"{prefix}-*.{ext}")) + sorted(outdir.glob(f"{prefix}.{ext}"))
    if cleanup_pdf:
        pdf.unlink(missing_ok=True)
    if not pages:
        return err("no images produced")
    total = 0
    for page in pages:
        total += page.stat().st_size
        print(f"page  {page}  ({page.stat().st_size:,} bytes)")
    print(f"\n{len(pages)} page image(s) at {args.dpi} dpi in {outdir}")
    if total > 8_000_000:
        print("NOTE: total image size is large; read only the pages you need, or re-render with --dpi 80.")
    return EXIT_OK


# ── new (markdown -> docx) ──────────────────────────────────────────────────


def cmd_new(args: argparse.Namespace) -> int:
    require_libs("new")
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import md2docx

    src = Path(args.markdown)
    if not src.exists():
        return err(f"no such file: {src}")
    out = Path(args.out) if args.out else src.with_suffix(".docx")
    md2docx.markdown_to_docx(
        src.read_text(encoding="utf-8"),
        out,
        template=Path(args.template) if args.template else None,
        title=args.title,
        toc=args.toc,
        page_numbers=args.page_numbers,
        landscape=args.landscape,
    )
    print(f"OK  {src}  ->  {out}  ({out.stat().st_size:,} bytes)")
    return EXIT_OK


# ── validate ────────────────────────────────────────────────────────────────


def cmd_validate(args: argparse.Namespace) -> int:
    """Structural + semantic checks. The point is catching silent breakage."""
    require_libs("validate")
    path = Path(args.file)
    if not path.exists():
        return err(f"no such file: {path}")
    if not zipfile.is_zipfile(path):
        return err(f"{path} is not a zip/OOXML package")

    problems: list[str] = []
    warnings: list[str] = []

    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            problems.append(f"corrupt zip member: {bad}")
        names = set(zf.namelist())
        if path.suffix.lower() == ".docx" and "word/document.xml" not in names:
            problems.append("word/document.xml missing")
        if "[Content_Types].xml" not in names:
            problems.append("[Content_Types].xml missing (Word will refuse to open it)")
        from lxml import etree

        for name in sorted(n for n in names if n.endswith(".xml") or n.endswith(".rels")):
            try:
                etree.fromstring(zf.read(name))
            except etree.XMLSyntaxError as exc:
                problems.append(f"malformed XML in {name}: {exc}")

        # Comments must have every cross-linked part, or Word reports the file
        # as unreadable rather than showing the comment.
        if any(n.startswith("word/comments") for n in names):
            for required in (
                "word/comments.xml",
                "word/_rels/document.xml.rels",
            ):
                if required not in names:
                    warnings.append(f"comments present but {required} missing")
            document = zf.read("word/document.xml").decode("utf-8", "replace")
            if "commentRangeStart" not in document:
                warnings.append("comments exist but no commentRangeStart marker anchors them to text")

        # Tracked changes: a deleted run must carry w:delText, not w:t.
        if path.suffix.lower() == ".docx" and "word/document.xml" in names:
            document = zf.read("word/document.xml").decode("utf-8", "replace")
            ins, dele = count_revisions(document)
            del_text = document.count("<w:delText")
            if dele and not del_text:
                warnings.append(f"{dele} <w:del> element(s) but no <w:delText>: deletions will show as unchanged text")
            if args.author:
                import re

                untracked = 0
                for match in re.finditer(r"<w:(?:ins|del)\b[^>]*w:author=\"([^\"]*)\"", document):
                    if match.group(1) != args.author:
                        untracked += 1
                if untracked:
                    warnings.append(f"{untracked} revision(s) authored by someone other than {args.author!r}")
            if args.original:
                problems.extend(compare_original(Path(args.original), document, args.author))

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        problems.extend(check_formulas(path))
    if path.suffix.lower() in {".docx", ".pptx", ".xlsx"}:
        proc = soffice("--convert-to", "pdf", "--outdir", str(path.parent / ".office-validate"), str(path))
        pdf = path.parent / ".office-validate" / (path.stem + ".pdf")
        if proc.returncode != 0 or not pdf.exists():
            problems.append("LibreOffice could not open the file (strong signal it is broken)")
        else:
            pdf.unlink(missing_ok=True)
            try:
                pdf.parent.rmdir()
            except OSError:
                pass

    for p in problems:
        print(f"PROBLEM  {p}")
    for w in warnings:
        print(f"warning  {w}")
    if problems:
        print(f"\nFAILED: {len(problems)} problem(s)")
        return EXIT_FAIL
    print(f"\nOK: {path.name} passed ({len(warnings)} warning(s))")
    return EXIT_OK


def check_formulas(path: Path) -> list[str]:
    import openpyxl

    problems: list[str] = []
    wb = openpyxl.load_workbook(str(path), data_only=True)
    error_literals = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!", "#N/A"}
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() in error_literals:
                    problems.append(f"formula error {cell.value} at {name}!{cell.coordinate}")
    if not problems:
        wb_f = openpyxl.load_workbook(str(path), data_only=False)
        empty_cache = 0
        for name in wb_f.sheetnames:
            for row in wb_f[name].iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        empty_cache += 1
        if empty_cache:
            pass  # cached values are absent until Excel/LibreOffice recalculates
    return problems


def accept_all_regex(xml: str) -> str:
    """XML equivalent of "accept all tracked changes".

    Deliberately NOT `<w:del\\b.*?</w:del>`: Word writes a self-closing `<w:del/>`
    inside `w:pPr` for a deleted paragraph mark, and that pattern then matches
    the tag against the *next* `</w:del>` - swallowing every element between
    them, including insertions. Stripping only the insertion wrappers and
    handling deletions at the extraction layer avoids the trap entirely.
    """
    import re

    return re.sub(r"<w:ins\b[^>]*>|</w:ins\s*>", "", xml)


# Matches `<w:t>` and `<w:t xml:space="preserve">` but NOT `<w:delText>`,
# `<w:instrText>`, or `<w:tab/>`. The trailing `(?:\s[^>]*)?` is load-bearing:
# a bare `[^>]*` happily consumes the `e` of `delText` and treats deleted text
# as document text - the single most consequential bug in this style of code.
W_T_RE = r"<w:t(?:\s[^>]*)?>(.*?)</w:t\s*>"


def extract_text(xml: str, include_deleted: bool = False) -> str:
    """Concatenated text of a document.xml fragment.

    `include_deleted=True` yields what a reviewer *sees* (the redline, with
    deletions still present); the default yields the accepted text.
    """
    import re

    text = "".join(re.findall(W_T_RE, xml, flags=re.DOTALL))
    if include_deleted:
        text += "".join(re.findall(r"<w:delText(?:\s[^>]*)?>(.*?)</w:delText\s*>", xml, flags=re.DOTALL))
    return re.sub(r"\s+", " ", text).strip()


def count_revisions(xml: str) -> tuple[int, int]:
    """Count insertions and deletions, without matching `<w:delText>` or `<w:instrText>`."""
    import re

    ins = len(re.findall(r"<w:ins\b", xml))
    dele = len(re.findall(r"<w:del(?:\s[^>]*)?\s*/?>", xml))
    return ins, dele


def compare_original(original: Path, edited_document_xml: str, author: str | None) -> list[str]:
    """Every text difference must sit inside a tracked revision.

    What this catches that a human reviewer cannot: an edit written directly
    into the XML looks identical, in the accepted view, to one made under a
    revision. The only reliable check is to strip the revision wrappers and
    compare the result against the original.
    """
    with zipfile.ZipFile(original) as zf:
        original_xml = zf.read("word/document.xml").decode("utf-8", "replace")
    before = extract_text(original_xml)
    as_shown = extract_text(edited_document_xml, include_deleted=True)
    as_accepted = extract_text(accept_all_regex(edited_document_xml))

    problems: list[str] = []
    if as_accepted != as_shown and not author:
        problems.append(
            "the document contains tracked revisions but no --author was given to verify them against"
        )
    if as_accepted != before and as_accepted == as_shown:
        problems.append(
            "text differs from the original but nothing is wrapped in a tracked revision: "
            "the edit would pass a reviewer's 'reject all' unnoticed"
        )
    if as_accepted == before and as_shown != before:
        problems.append(
            "the accepted text is unchanged, so the revisions cancel out: either the edit was "
            "reverted or the same text was marked as both inserted and deleted"
        )
    return problems


# ── main ────────────────────────────────────────────────────────────────────


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="office.py",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = ap.add_subparsers(dest="command", required=True)

    p = sub.add_parser("info", help="identify a file and summarize its package")
    p.add_argument("file")
    p.add_argument("--json", action="store_true", help="also print machine-readable JSON")
    p.set_defaults(func=cmd_info)

    p = sub.add_parser("dump", help="extract content as structured text (no LibreOffice needed)")
    p.add_argument("file")
    p.add_argument("--out", help="write to this path instead of stdout")
    p.add_argument("--max-rows", type=int, default=200, help="xlsx row cap (default 200)")
    p.add_argument("--max-cols", type=int, default=40, help="xlsx column cap (default 40)")
    p.set_defaults(func=cmd_dump)

    p = sub.add_parser("convert", help="convert between formats with LibreOffice")
    p.add_argument("file")
    p.add_argument("--to", required=True, help=f"target format: {', '.join(sorted(CONVERT_TARGETS))}")
    p.add_argument("--outdir", help="output directory (default: alongside the source)")
    p.set_defaults(func=cmd_convert)

    p = sub.add_parser("render", help="rasterize pages to images for visual review")
    p.add_argument("file")
    p.add_argument("--outdir", help="output directory (default: --prefix directory)")
    p.add_argument("--prefix", default="./page", help="output basename (default: ./page)")
    p.add_argument("--dpi", type=int, default=110, help="raster dpi (default 110; use 80 to save context)")
    p.add_argument("--pages", help="page range, e.g. 1 or 2-4")
    p.add_argument("--format", default="png", choices=["png", "jpeg"])
    p.add_argument("--keep-pdf", action="store_true", help="keep the intermediate PDF")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("new", help="build a .docx from a Markdown file")
    p.add_argument("markdown")
    p.add_argument("--out", help="output .docx (default: alongside the Markdown)")
    p.add_argument("--template", help="existing .docx to inherit styles and branding from")
    p.add_argument("--title", help="document title for the properties and cover heading")
    p.add_argument("--toc", action="store_true", help="insert a table-of-contents field after the title")
    p.add_argument("--page-numbers", action="store_true", help="add a footer with page numbers")
    p.add_argument("--landscape", action="store_true", help="landscape page orientation")
    p.set_defaults(func=cmd_new)

    p = sub.add_parser("validate", help="structural and semantic checks before delivering a file")
    p.add_argument("file")
    p.add_argument("--original", help="for edited documents: the pre-edit file, to catch untracked edits")
    p.add_argument("--author", help="the revision author name every tracked edit should carry")
    p.set_defaults(func=cmd_validate)

    return ap


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
