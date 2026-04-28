"""Dump every page of a OneNote notebook to its 2013 XML form on disk.

Talks to OneNote desktop via the in-GAC ``Microsoft.Office.Interop.OneNote``
assembly (loaded through ``pythonnet``), so it bypasses the late-bound
IDispatch path that Office click-to-run installs leave broken on many
machines.

OneNote desktop must already be running with the target notebook open.
"""

from __future__ import annotations

import hashlib
import re
import sys
import time
from pathlib import Path
from xml.etree import ElementTree as ET


_NS = {"on": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


def _sanitize(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.strip().rstrip(".")
    return name[:80] if len(name) > 80 else name


def dump_notebook(notebook_name: str, out_root: Path) -> tuple[int, int]:
    """Dump every page of ``notebook_name`` under ``out_root``.

    Returns ``(pages_dumped, pages_errored)``.
    """
    import clr  # type: ignore  # provided by pythonnet
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
        Application, HierarchyScope, PageInfo, XMLSchema,
    )

    out_root = Path(out_root)
    out_root.mkdir(parents=True, exist_ok=True)

    print("Connecting to OneNote ...", flush=True)
    app = Application()

    hierarchy_xml = app.GetHierarchy(None, HierarchyScope.hsPages, "")
    (out_root / "_hierarchy.xml").write_text(hierarchy_xml, encoding="utf-8")
    print(f"Hierarchy saved: {len(hierarchy_xml)} chars", flush=True)

    root = ET.fromstring(hierarchy_xml)
    target = next(
        (nb for nb in root.findall("on:Notebook", _NS) if nb.attrib.get("name") == notebook_name),
        None,
    )
    if target is None:
        names = [nb.attrib.get("name") for nb in root.findall("on:Notebook", _NS)]
        raise RuntimeError(
            f"Notebook {notebook_name!r} not found. Open notebooks: {names}"
        )

    pages_done = 0
    pages_err = 0
    skip_log: list[tuple[str, str, str]] = []

    def walk(parent: ET.Element, parts: list[str]) -> None:
        nonlocal pages_done, pages_err
        for child in parent:
            tag = child.tag.split("}", 1)[-1]
            name = child.attrib.get("name", tag)
            if tag in ("SectionGroup", "Section"):
                walk(child, parts + [_sanitize(name)])
            elif tag == "Page":
                page_id = child.attrib["ID"]
                section_dir = out_root.joinpath(*parts) if parts else out_root
                section_dir.mkdir(parents=True, exist_ok=True)
                fname = (
                    _sanitize(name)
                    + "_"
                    + hashlib.md5(page_id.encode()).hexdigest()[:8]
                    + ".xml"
                )
                page_path = section_dir / fname
                if page_path.exists() and page_path.stat().st_size > 0:
                    pages_done += 1
                    continue
                try:
                    page_xml = app.GetPageContent(
                        page_id, "", PageInfo.piBasic, XMLSchema.xs2013
                    )
                    page_path.write_text(page_xml, encoding="utf-8")
                    pages_done += 1
                    if pages_done % 25 == 0:
                        print(f"  ... {pages_done} pages", flush=True)
                    time.sleep(0.05)
                except Exception as exc:  # noqa: BLE001
                    pages_err += 1
                    skip_log.append((name, page_id, str(exc)))
                    print(f"  ERR {name}: {exc}", file=sys.stderr, flush=True)
                    time.sleep(0.5)

    walk(target, [_sanitize(notebook_name)])

    if skip_log:
        with (out_root / "_errors.log").open("w", encoding="utf-8") as f:
            for n, i, e in skip_log:
                f.write(f"{n}\t{i}\t{e}\n")

    print(f"Done. Pages dumped: {pages_done}, errors: {pages_err}", flush=True)
    return pages_done, pages_err
