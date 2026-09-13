#!/usr/bin/env python3
"""docx_edit - safe, reviewable edits to an existing .docx.

Editing a Word file is where naive approaches quietly corrupt documents. The
three traps this module exists to avoid:

1. **Split runs.** Word fragments a sentence across many `<w:r>` elements for
   revision ids, spell-check markers, and formatting boundaries. The phrase you
   can see is usually not a contiguous string in the XML, so a find-and-replace
   over `paragraph.text` silently misses it. `merge_runs` coalesces adjacent
   identically-formatted runs first.

2. **Untracked edits.** Under a redline workflow every change must be wrapped in
   `<w:ins>`/`<w:del>`. A plain text mutation is invisible in the accepted view
   and survives a reviewer's "reject all".

3. **Broken comment parts.** A comment lives in six cross-linked package parts.
   Writing one file by hand yields a document Word refuses to open, or a comment
   with no anchor to any text.

Everything here operates on XML trees and never pretty-prints, because
reformatting the XML changes whitespace inside `<w:t xml:space="preserve">` and
alters the rendered text.

Usage:
    docx_edit.py find-replace in.docx out.docx --old "ACME" --new "Acme Corp" \
        --track --author "A. Reviewer"
    docx_edit.py merge-runs in.docx -o merged.docx
    docx_edit.py comment contract.docx -o annotated.docx --text "Cap is too low"
    docx_edit.py accept in.docx out.docx          # accept all tracked changes
    docx_edit.py reject in.docx out.docx          # reject all tracked changes
"""

from __future__ import annotations

import argparse
import datetime as _dt
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"

NSMAP = {"w": W, "r": R}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def rq(tag: str) -> str:
    return f"{{{R}}}{tag}"


# ── package read/write ──────────────────────────────────────────────────────


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def write_parts(parts: dict[str, bytes], out: Path) -> None:
    """Rewrite the package deterministically.

    `[Content_Types].xml` must stay the first entry, and deflate keeps files
    smaller than a Word re-save would. Symlink entries are dropped: a .docx from
    an external party is untrusted input.
    """
    ordered = ["[Content_Types].xml", "_rels/.rels"] + sorted(
        n for n in parts if n not in {"[Content_Types].xml", "_rels/.rels"}
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in ordered:
            if name in parts:
                zf.writestr(name, parts[name])


def parse(parts: dict[str, bytes], name: str) -> etree._Element | None:
    if name not in parts:
        return None
    return etree.fromstring(parts[name])


# ── run coalescing ──────────────────────────────────────────────────────────


def _rpr_key(run: etree._Element) -> bytes | None:
    rpr = run.find(q("rPr"))
    return None if rpr is None else etree.tostring(rpr)


def merge_runs_in_tree(tree: etree._Element) -> int:
    """Merge adjacent identically-formatted runs, so visible text is findable.

    Content and rendering are unchanged: only run boundaries move. Runs inside
    `<w:ins>`/`<w:del>` are merged only with runs in the same revision, so
    tracked edits are never reshaped.
    """
    merged = 0
    for parent in tree.iter():
        if parent.tag not in (q("p"), q("hyperlink"), q("ins"), q("del"), q("smartTag"), q("sdtContent")):
            continue
        children = list(parent)
        idx = 0
        while idx < len(children) - 1:
            current, nxt = children[idx], children[idx + 1]
            if current.tag != q("r") or nxt.tag != q("r"):
                idx += 1
                continue
            if _rpr_key(current) != _rpr_key(nxt):
                idx += 1
                continue
            cur_texts = current.findall(q("t")) + current.findall(q("delText"))
            nxt_texts = nxt.findall(q("t")) + nxt.findall(q("delText"))
            if len(cur_texts) != 1 or len(nxt_texts) != 1:
                idx += 1
                continue
            if not cur_texts[0].text or not nxt_texts[0].text:
                idx += 1
                continue
            cur_texts[0].text = cur_texts[0].text + nxt_texts[0].text
            parent.remove(nxt)
            merged += 1
            children = list(parent)
        # descend handled by tree.iter()
    return merged


# ── tracked find & replace ──────────────────────────────────────────────────


class RevisionIds:
    """Monotonic w:id allocator that cannot collide with existing revisions."""

    def __init__(self, tree: etree._Element | None) -> None:
        self.next_id = 9000
        if tree is None:
            return
        for el in tree.iter():
            if el.tag in (q("ins"), q("del")) and el.get(q("id")):
                try:
                    self.next_id = max(self.next_id, int(el.get(q("id"))) + 1)
                except ValueError:
                    pass

    def take(self) -> str:
        value = str(self.next_id)
        self.next_id += 1
        return value


def _timestamp() -> str:
    return _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _text_element(run: etree._Element) -> etree._Element | None:
    texts = run.findall(q("t"))
    return texts[0] if len(texts) == 1 else None


def _make_run(template: etree._Element, text: str, deleted: bool) -> etree._Element:
    """Clone `template`'s formatting into a run holding `text`."""
    run = etree.Element(q("r"))
    rpr = template.find(q("rPr"))
    if rpr is not None:
        run.append(etree.fromstring(etree.tostring(rpr)))
    el = etree.SubElement(run, q("delText") if deleted else q("t"))
    el.set(f"{{{XML}}}space", "preserve")
    el.text = text
    return run


def replace_in_paragraph(
    paragraph: etree._Element,
    pattern: re.Pattern[str],
    replacement: str,
    track: bool,
    author: str,
    ids: RevisionIds,
) -> int:
    """Replace `pattern` inside one paragraph; returns the number of hits.

    With `track`, each hit becomes `<w:del>` (old text as `w:delText`) followed
    by `<w:ins>` (new text) at the hit position, so the paragraph's other runs
    keep their order and formatting.
    """
    hits = 0
    runs = [r for r in paragraph.findall(q("r")) if _has_visible_text(r)]
    for run in runs:
        text_el = _text_element(run)
        if text_el is None or not text_el.text:
            continue
        text = text_el.text
        if not pattern.search(text):
            continue
        pieces: list[etree._Element] = []
        cursor = 0
        for match in pattern.finditer(text):
            hits += 1
            if match.start() > cursor:
                pieces.append(_make_run(run, text[cursor : match.start()], deleted=False))
            if track:
                deletion = etree.Element(q("del"))
                deletion.set(q("id"), ids.take())
                deletion.set(q("author"), author)
                deletion.set(q("date"), _timestamp())
                deletion.append(_make_run(run, match.group(0), deleted=True))
                pieces.append(deletion)
                if replacement:
                    insertion = etree.Element(q("ins"))
                    insertion.set(q("id"), ids.take())
                    insertion.set(q("author"), author)
                    insertion.set(q("date"), _timestamp())
                    insertion.append(_make_run(run, replacement, deleted=False))
                    pieces.append(insertion)
            elif replacement:
                pieces.append(_make_run(run, replacement, deleted=False))
            cursor = match.end()
        if cursor < len(text):
            pieces.append(_make_run(run, text[cursor:], deleted=False))
        parent = run.getparent()
        index = list(parent).index(run)
        parent.remove(run)
        for offset, piece in enumerate(pieces):
            parent.insert(index + offset, piece)
    return hits


def _has_visible_text(run: etree._Element) -> bool:
    return bool(run.findall(q("t"))) or bool(run.findall(q("delText")))


def find_replace(
    src: Path,
    out: Path,
    old: str,
    new: str,
    track: bool = True,
    author: str = "DSH",
    regex: bool = False,
    dry_run: bool = False,
) -> dict:
    parts = read_parts(src)
    document = parse(parts, "word/document.xml")
    if document is None:
        raise SystemExit(f"ERROR: {src} has no word/document.xml")
    merged = merge_runs_in_tree(document)
    ids = RevisionIds(document)
    pattern = re.compile(old if regex else re.escape(old))
    total = 0
    for paragraph in document.iter(q("p")):
        total += replace_in_paragraph(paragraph, pattern, new, track, author, ids)
    if total == 0:
        return {"hits": 0, "merged_runs": merged, "written": False}
    if dry_run:
        return {"hits": total, "merged_runs": merged, "written": False}
    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_parts(parts, out)
    return {"hits": total, "merged_runs": merged, "written": True}


# ── comments ────────────────────────────────────────────────────────────────

COMMENTS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{W}"/>
"""

COMMENTS_EXTENDED_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>
"""

COMMENTS_IDS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>
"""


def _next_rel_id(rels: etree._Element) -> str:
    used = {rel.get("Id") for rel in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    return f"rId{n}"


def add_comment(
    src: Path,
    out: Path,
    text: str,
    author: str,
    anchor: str | None,
    parent: int | None,
    initials: str,
) -> dict:
    """Attach a comment, anchored to `anchor` text when given.

    All cross-linked parts are created or updated: the comment body, the
    relationship, the content-type override, and - when an anchor is provided -
    the commentRangeStart/End/Reference markers in `document.xml` without which
    the comment exists but is invisible.
    """
    parts = read_parts(src)
    document = parse(parts, "word/document.xml")
    if document is None:
        raise SystemExit(f"ERROR: {src} has no word/document.xml")

    # Body part: create or extend.
    if "word/comments.xml" in parts:
        comments = parse(parts, "word/comments.xml")
    else:
        comments = etree.fromstring(COMMENTS_XML.encode())
    existing = [c for c in comments.findall(q("comment"))]
    comment_id = max([int(c.get(q("id"), "0")) for c in existing], default=-1) + 1

    comment = etree.SubElement(comments, q("comment"))
    comment.set(q("id"), str(comment_id))
    comment.set(q("author"), author)
    comment.set(q("date"), _timestamp())
    comment.set(q("initials"), initials)
    if parent is not None:
        comment.set(q("parent"), str(parent))
    para = etree.SubElement(comment, q("p"))
    run = etree.SubElement(para, q("r"))
    t = etree.SubElement(run, q("t"))
    t.set(f"{{{XML}}}space", "preserve")
    t.text = text
    parts["word/comments.xml"] = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Relationship + content type.
    rels_name = "word/_rels/document.xml.rels"
    rels = parse(parts, rels_name)
    if rels is None:
        rels = etree.fromstring(
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PR}"/>'.encode()
        )
    rel_id = None
    for rel in rels:
        if rel.get("Target") in {"comments.xml", "/word/comments.xml"}:
            rel_id = rel.get("Id")
            break
    if rel_id is None:
        rel_id = _next_rel_id(rels)
        rel = etree.SubElement(rels, f"{{{PR}}}Relationship")
        rel.set("Id", rel_id)
        rel.set("Type", f"{R}/comments")
        rel.set("Target", "comments.xml")
        parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    else:
        parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    content_types = parse(parts, "[Content_Types].xml")
    has_override = any(
        el.get("PartName") == "/word/comments.xml" for el in content_types.findall(f"{{{CT}}}Override")
    )
    if not has_override:
        override = etree.SubElement(content_types, f"{{{CT}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )
        parts["[Content_Types].xml"] = etree.tostring(
            content_types, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    anchored = False
    if anchor:
        merged = merge_runs_in_tree(document)
        anchored = _anchor_comment(document, anchor, comment_id, author, initials)
        if not anchored:
            raise SystemExit(
                f"ERROR: anchor text {anchor!r} not found in the document; the comment would be invisible.\n"
                "       Pass --anchor with text that exists, or omit it to create an unanchored comment."
            )
    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_parts(parts, out)
    return {"comment_id": comment_id, "anchored": anchored, "written": True}


def _anchor_comment(document: etree._Element, anchor: str, comment_id: int, author: str, initials: str) -> bool:
    """Wrap the first occurrence of `anchor` in comment range markers."""
    for paragraph in document.iter(q("p")):
        runs = [r for r in paragraph.findall(q("r")) if _has_visible_text(r)]
        for run in runs:
            text_el = _text_element(run)
            if text_el is None or not text_el.text or anchor not in text_el.text:
                continue
            text = text_el.text
            start_idx = text.index(anchor)
            end_idx = start_idx + len(anchor)
            fragment = text[start_idx:end_idx]
            parent = run.getparent()
            position = list(parent).index(run)

            def marker(tag: str, with_ref: bool) -> etree._Element:
                el = etree.Element(q(tag))
                el.set(q("id"), str(comment_id))
                if with_ref:
                    el.set(q("author"), author)
                    el.set(q("date"), _timestamp())
                    el.set(q("initials"), initials)
                return el

            replacement: list[etree._Element] = []
            if start_idx > 0:
                replacement.append(_make_run(run, text[:start_idx], deleted=False))
            replacement.append(marker("commentRangeStart", with_ref=False))
            replacement.append(_make_run(run, fragment, deleted=False))
            replacement.append(marker("commentRangeEnd", with_ref=False))
            ref_run = etree.Element(q("r"))
            rpr = run.find(q("rPr"))
            if rpr is not None:
                ref_run.append(etree.fromstring(etree.tostring(rpr)))
            ref = etree.SubElement(ref_run, q("commentReference"))
            ref.set(q("id"), str(comment_id))
            replacement.append(ref_run)
            if end_idx < len(text):
                replacement.append(_make_run(run, text[end_idx:], deleted=False))

            parent.remove(run)
            for offset, el in enumerate(replacement):
                parent.insert(position + offset, el)
            return True
    return False


# ── accept / reject revisions ───────────────────────────────────────────────


def resolve_revisions(src: Path, out: Path, accept: bool) -> dict:
    """Accept or reject all tracked changes, joining deleted paragraph marks."""
    parts = read_parts(src)
    document = parse(parts, "word/document.xml")
    if document is None:
        raise SystemExit(f"ERROR: {src} has no word/document.xml")
    counts = {"insertions": 0, "deletions": 0, "paragraph_marks": 0}

    for ins in list(document.iter(q("ins"))):
        counts["insertions"] += 1
        parent = ins.getparent()
        index = list(parent).index(ins)
        parent.remove(ins)
        if accept:
            for offset, child in enumerate(list(ins)):
                parent.insert(index + offset, child)

    for dele in list(document.iter(q("del"))):
        parent = dele.getparent()
        if parent.tag == q("rPr"):
            counts["paragraph_marks"] += 1
            parent.remove(dele)
            if accept:
                _mark_paragraph_deleted(parent)
            continue
        counts["deletions"] += 1
        index = list(parent).index(dele)
        parent.remove(dele)
        if not accept:
            for offset, child in enumerate(list(dele)):
                if child.tag == q("r"):
                    for del_text in child.findall(q("delText")):
                        del_text.tag = q("t")
                parent.insert(index + offset, child)

    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_parts(parts, out)
    return counts


def _mark_paragraph_deleted(rpr: etree._Element) -> None:
    """Record an accepted paragraph deletion as `w:rPr/w:del` on the mark."""
    existing = rpr.find(q("del"))
    if existing is None:
        el = etree.Element(q("del"))
        el.set(q("id"), "0")
        el.set(q("author"), "accept")
        el.set(q("date"), _timestamp())
        rpr.insert(0, el)


# ── CLI ─────────────────────────────────────────────────────────────────────


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="command", required=True)

    p = sub.add_parser("find-replace", help="replace text, optionally as tracked revisions")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--old", required=True)
    p.add_argument("--new", default="")
    p.add_argument("--regex", action="store_true", help="treat --old as a regular expression")
    p.add_argument("--no-track", action="store_true", help="edit directly instead of tracked")
    p.add_argument("--author", default="DSH", help="revision author recorded in the document")
    p.add_argument("--dry-run", action="store_true", help="report hit count without writing")

    p = sub.add_parser("merge-runs", help="coalesce fragmented runs so text is findable")
    p.add_argument("src")
    p.add_argument("-o", "--out", required=True)

    p = sub.add_parser("comment", help="add a comment, anchored to specific text")
    p.add_argument("src")
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--text", required=True, help="the comment body")
    p.add_argument("--anchor", help="existing document text to attach the comment to")
    p.add_argument("--parent", type=int, help="id of the comment this one replies to")
    p.add_argument("--author", default="DSH")
    p.add_argument("--initials", default="DSH")

    p = sub.add_parser("accept", help="accept every tracked change")
    p.add_argument("src")
    p.add_argument("out")

    p = sub.add_parser("reject", help="reject every tracked change")
    p.add_argument("src")
    p.add_argument("out")

    args = ap.parse_args(argv)
    src = Path(args.src)
    if not src.exists():
        print(f"ERROR: no such file: {src}")
        return 1

    if args.command == "find-replace":
        result = find_replace(
            src,
            Path(args.out),
            args.old,
            args.new,
            track=not args.no_track,
            author=args.author,
            regex=args.regex,
            dry_run=args.dry_run,
        )
        print(
            f"{'would replace' if args.dry_run else 'replaced'} {result['hits']} occurrence(s); "
            f"merged {result['merged_runs']} fragmented run(s)"
        )
        if result["hits"] == 0:
            print("No match. The text may be split across runs; this command merges runs first, so the\n"
                  "remaining causes are different text (whitespace, curly quotes) or text inside a table/textbox.")
            return 1
        if result["written"]:
            print(f"wrote {args.out}")
        return 0

    if args.command == "merge-runs":
        parts = read_parts(src)
        document = parse(parts, "word/document.xml")
        merged = merge_runs_in_tree(document)
        parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
        write_parts(parts, Path(args.out))
        print(f"merged {merged} run(s); wrote {args.out}")
        return 0

    if args.command == "comment":
        result = add_comment(
            src,
            Path(args.out),
            args.text,
            args.author,
            args.anchor,
            args.parent,
            args.initials,
        )
        print(f"comment id {result['comment_id']} ({'anchored' if result['anchored'] else 'unanchored'}); wrote {args.out}")
        return 0

    counts = resolve_revisions(src, Path(args.out), accept=args.command == "accept")
    print(
        f"{args.command}ed {counts['insertions']} insertion(s), {counts['deletions']} deletion(s), "
        f"{counts['paragraph_marks']} paragraph mark(s); wrote {args.out}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
