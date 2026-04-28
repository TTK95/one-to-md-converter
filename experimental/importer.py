"""Send a cleaned Markdown page back into OneNote via UpdatePageContent.

Resolves the target OneNote page by matching the 8-char filename hash that
``one_to_md.dumper`` writes (``<title>_<md5(id)[:8]>.md``). Converts MD to
OneNote 2013 XML via :mod:`md_to_xml`, then calls
``IApplication.UpdatePageContent``.

Usage:
    python importer.py <NotebookName> <markdown_file> [--dry-run] [--debug-xml PATH]

Prerequisite:
    OneNote desktop running with the target notebook open.

Phase 1 vs Phase 2:
    If the markdown lacks oid annotations, ``UpdatePageContent`` *adds* content
    rather than replacing -- only safe against an empty sandbox page. With oid
    annotations (default of ``one-to-md convert``), OneNote merges by
    ``objectID`` and edits replace matching blocks.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))


NS = {"on": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


def _hash_id(page_id: str) -> str:
    return hashlib.md5(page_id.encode()).hexdigest()[:8]


def find_page_id(
    hierarchy_xml: str, notebook_name: str, hash_suffix: str
) -> tuple[str, str, str] | None:
    """Walk the hierarchy and return the page ID whose md5(id)[:8] matches.

    Raises ``RuntimeError`` if more than one page in the notebook hashes to the
    same suffix -- cosmetically rare for an 8-char hash on a single notebook,
    but worth detecting before silently picking the wrong page. Returns
    ``(page_id, page_name, lastModifiedTime)`` or ``None`` if no match.
    """
    root = ET.fromstring(hierarchy_xml)
    target_nb = next(
        (
            nb
            for nb in root.findall("on:Notebook", NS)
            if nb.attrib.get("name") == notebook_name
        ),
        None,
    )
    if target_nb is None:
        return None

    matches: list[tuple[str, str, str]] = []

    def walk(elem: ET.Element) -> None:
        for child in elem:
            tag = child.tag.split("}", 1)[-1]
            if tag in ("SectionGroup", "Section"):
                walk(child)
            elif tag == "Page":
                pid = child.attrib.get("ID", "")
                if _hash_id(pid) == hash_suffix:
                    matches.append(
                        (
                            pid,
                            child.attrib.get("name", ""),
                            child.attrib.get("lastModifiedTime", ""),
                        )
                    )

    walk(target_nb)

    if not matches:
        return None
    if len(matches) > 1:
        details = "\n  ".join(f"{pid}  ({name})" for pid, name, _ in matches)
        raise RuntimeError(
            f"Hash collision: {len(matches)} pages in {notebook_name!r} hash to "
            f"{hash_suffix!r}. Disambiguate by widening the hash. Candidates:\n  {details}"
        )
    return matches[0]


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        prog="importer",
        description=(
            "Push a Markdown file (as produced by `one-to-md convert`) back "
            "into the matching OneNote page via UpdatePageContent."
        ),
    )
    p.add_argument("notebook", help="Notebook display name (as shown in OneNote).")
    p.add_argument("markdown", type=Path, help="Path to the .md file.")
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Build the XML and report the target page, but do not call UpdatePageContent.",
    )
    p.add_argument(
        "--debug-xml",
        type=Path,
        default=None,
        help="Write the generated OneNote XML to this path before sending. Useful for diffing schema rejections.",
    )
    args = p.parse_args(argv)

    md_path: Path = args.markdown
    if not md_path.is_file():
        print(f"Not a file: {md_path}", file=sys.stderr)
        return 2

    m = re.match(r"^.+_([0-9a-f]{8})$", md_path.stem)
    if not m:
        print(
            f"Filename must end with _<8charhash>.md (the hash dumper.py writes); got {md_path.name}",
            file=sys.stderr,
        )
        return 2
    hash_suffix = m.group(1)

    try:
        from md_to_xml import md_to_xml  # type: ignore  # sibling import
    except ImportError as exc:
        print(
            f"md_to_xml not importable: {exc}\n"
            f"Install markdown-it-py and run from the experimental/ directory.",
            file=sys.stderr,
        )
        return 4

    try:
        import clr  # type: ignore  # provided by pythonnet
    except ImportError as exc:
        print(f"pythonnet not installed: {exc}", file=sys.stderr)
        return 4

    try:
        clr.AddReference("Microsoft.Office.Interop.OneNote")
    except Exception:
        import glob

        candidates = sorted(
            glob.glob(
                r"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\*\Microsoft.Office.Interop.OneNote.dll"
            ),
            reverse=True,
        )
        if not candidates:
            raise
        clr.AddReference(candidates[0])

    from Microsoft.Office.Interop.OneNote import (  # type: ignore
        Application,
        HierarchyScope,
        XMLSchema,
    )
    from System import DateTime  # type: ignore

    print("Connecting to OneNote ...", flush=True)
    app = Application()

    print(
        f"Resolving page for hash {hash_suffix} in {args.notebook!r} ...", flush=True
    )
    hierarchy_xml = app.GetHierarchy(None, HierarchyScope.hsPages, "")
    found = find_page_id(hierarchy_xml, args.notebook, hash_suffix)
    if found is None:
        print(
            f"No page with hash {hash_suffix} in notebook {args.notebook!r}.",
            file=sys.stderr,
        )
        return 3
    page_id, page_name, last_mod = found
    print(f"  page: {page_name!r}  ID={page_id}  lastModified={last_mod}", flush=True)

    md_text = md_path.read_text(encoding="utf-8")
    page_xml = md_to_xml(md_text, page_id=page_id)

    if args.debug_xml:
        args.debug_xml.parent.mkdir(parents=True, exist_ok=True)
        args.debug_xml.write_text(page_xml, encoding="utf-8")
        print(f"Wrote XML to {args.debug_xml}", flush=True)

    if args.dry_run:
        print(f"[dry-run] Would call UpdatePageContent ({len(page_xml)} chars).")
        return 0

    print(f"Sending UpdatePageContent ({len(page_xml)} chars XML) ...", flush=True)
    # 4th arg force=True so dateExpectedLastModified isn't enforced.
    app.UpdatePageContent(page_xml, DateTime.MinValue, XMLSchema.xs2013, True)
    print("OK.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
