"""Send a cleaned Markdown page back into OneNote via UpdatePageContent.

Resolves the target OneNote page by matching the 8-char filename hash that
dumper.py writes (`<title>_<md5(id)[:8]>.md`). Converts MD -> OneNote 2013 XML
via md_to_xml, then calls IApplication.UpdatePageContent.

Usage:
    python importer.py <NotebookName> <markdown_file>

Prerequisite:
    OneNote desktop running with the target notebook open.

WARNING (Phase 1 importer):
    UpdatePageContent merges by objectID. This script sends a *new* outline
    without objectIDs, which means OneNote ADDS content rather than replacing.
    For a clean test, point this at a freshly-created EMPTY page in a sandbox
    section ("_Test Restructure"). Do not run against production pages until
    a merge-aware Phase 2 is built.
"""

import re
import sys
import hashlib
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent))
from md_to_xml import md_to_xml

NS = {"on": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


def _hash_id(page_id: str) -> str:
    return hashlib.md5(page_id.encode()).hexdigest()[:8]


def find_page_id(hierarchy_xml: str, notebook_name: str, hash_suffix: str):
    """Walk the hierarchy and return the page ID whose md5(id)[:8] matches.

    Returns (page_id, page_name, lastModifiedTime) or None.
    """
    root = ET.fromstring(hierarchy_xml)
    target_nb = None
    for nb in root.findall("on:Notebook", NS):
        if nb.attrib.get("name") == notebook_name:
            target_nb = nb
            break
    if target_nb is None:
        return None

    def walk(elem):
        for child in elem:
            tag = child.tag.split("}", 1)[-1]
            if tag in ("SectionGroup", "Section"):
                r = walk(child)
                if r is not None:
                    return r
            elif tag == "Page":
                pid = child.attrib.get("ID", "")
                if _hash_id(pid) == hash_suffix:
                    return (
                        pid,
                        child.attrib.get("name", ""),
                        child.attrib.get("lastModifiedTime", ""),
                    )
        return None

    return walk(target_nb)


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    notebook_name = sys.argv[1]
    md_path = Path(sys.argv[2])
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

    import clr  # type: ignore
    clr.AddReference("Microsoft.Office.Interop.OneNote")
    from Microsoft.Office.Interop.OneNote import (  # type: ignore
        Application, HierarchyScope, XMLSchema,
    )
    from System import DateTime  # type: ignore

    print("Connecting to OneNote ...", flush=True)
    app = Application()

    print(f"Resolving page for hash {hash_suffix} in '{notebook_name}' ...", flush=True)
    hierarchy_xml = app.GetHierarchy(None, HierarchyScope.hsPages, "")
    found = find_page_id(hierarchy_xml, notebook_name, hash_suffix)
    if found is None:
        print(
            f"No page with hash {hash_suffix} in notebook '{notebook_name}'.",
            file=sys.stderr,
        )
        return 3
    page_id, page_name, last_mod = found
    print(f"  page: {page_name!r}  ID={page_id}", flush=True)

    md_text = md_path.read_text(encoding="utf-8")
    page_xml = md_to_xml(md_text, page_id=page_id)

    # Pass force=True (4th arg) so the dateExpectedLastModified isn't enforced.
    print(f"Sending UpdatePageContent ({len(page_xml)} chars XML) ...", flush=True)
    app.UpdatePageContent(page_xml, DateTime.MinValue, XMLSchema.xs2013, True)
    print("OK.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
