"""Convert OneNote 2013 XML page files (as produced by :mod:`one_to_md.dumper`)
into Markdown. Walks an input tree of XML and writes a parallel tree of ``.md``.
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET


_NS = {"on": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


def _render_t(text_html: str) -> str:
    """Render a ``<one:T>`` CDATA payload into Markdown inline text.

    OneNote stores inline formatting as ``<span style="...">`` plus a few
    standard tags. We preserve bold / italic / links and drop the rest.
    """
    if not text_html:
        return ""
    s = text_html

    def replace_span(m: re.Match) -> str:
        style = (m.group(1) or "").replace(" ", "")
        inner = m.group(2)
        if "font-weight:bold" in style:
            return f"**{inner}**"
        if "font-style:italic" in style:
            return f"*{inner}*"
        return inner

    s = re.sub(r'<span\s+style="([^"]*)">(.*?)</span>', replace_span, s, flags=re.DOTALL)
    s = re.sub(r"<span[^>]*>(.*?)</span>", r"\1", s, flags=re.DOTALL)
    s = re.sub(r"<b>(.*?)</b>", r"**\1**", s, flags=re.DOTALL)
    s = re.sub(r"<i>(.*?)</i>", r"*\1*", s, flags=re.DOTALL)
    s = re.sub(r"<u>(.*?)</u>", r"\1", s, flags=re.DOTALL)
    s = re.sub(r'<a\s+href="([^"]+)">(.*?)</a>', r"[\2](\1)", s, flags=re.DOTALL)
    s = s.replace("<br>", "  \n").replace("<br/>", "  \n").replace("<br />", "  \n")
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s)


def _render_oe(oe: ET.Element, indent: int = 0) -> list[str]:
    """Render a OneNote outline element (and children) to Markdown lines."""
    out: list[str] = []
    is_bullet = oe.find("on:List/on:Bullet", _NS) is not None
    is_numbered = oe.find("on:List/on:Number", _NS) is not None
    pad = "  " * indent

    text_parts = [_render_t(t.text) for t in oe.findall("on:T", _NS) if t.text]
    text = " ".join(p for p in text_parts if p).strip()

    if text:
        if is_bullet:
            out.append(f"{pad}- {text}")
        elif is_numbered:
            out.append(f"{pad}1. {text}")
        else:
            out.append(f"{pad}{text}")
            out.append("")
    else:
        img = oe.find("on:Image", _NS)
        if img is not None:
            out.append(f"{pad}![{img.attrib.get('alt', 'image')}](image)")
        ifile = oe.find("on:InsertedFile", _NS)
        if ifile is not None:
            out.append(f"{pad}_(attached file: {ifile.attrib.get('preferredName', 'file')})_")

    children = oe.find("on:OEChildren", _NS)
    if children is not None:
        nested_indent = indent + (1 if (is_bullet or is_numbered) else 0)
        for child_oe in children.findall("on:OE", _NS):
            out.extend(_render_oe(child_oe, nested_indent))

    return out


def _render_table(tbl: ET.Element) -> list[str]:
    rows: list[list[str]] = []
    for row in tbl.findall("on:Row", _NS):
        cells: list[str] = []
        for cell in row.findall("on:Cell", _NS):
            cell_lines: list[str] = []
            for oechildren in cell.findall("on:OEChildren", _NS):
                for oe in oechildren.findall("on:OE", _NS):
                    cell_lines.extend(_render_oe(oe))
            cells.append(" ".join(l.strip() for l in cell_lines).strip() or " ")
        rows.append(cells)
    if not rows:
        return []
    n = max(len(r) for r in rows)
    rows = [r + [" "] * (n - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * n) + " |"]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return out


def render_page(xml_text: str) -> str:
    """Convert a single OneNote 2013 ``<one:Page>`` document to Markdown."""
    root = ET.fromstring(xml_text)
    name = root.attrib.get("name", "Untitled")
    dt = root.attrib.get("dateTime", "")

    lines: list[str] = [f"# {name}", ""]
    if dt:
        lines.append(f"_Created: {dt}_")
        lines.append("")

    for child in root:
        tag = child.tag.split("}", 1)[-1]
        if tag != "Outline":
            continue
        for oechildren in child.findall("on:OEChildren", _NS):
            for oe in oechildren.findall("on:OE", _NS):
                table = oe.find("on:Table", _NS)
                if table is not None:
                    lines.extend(_render_table(table))
                    lines.append("")
                else:
                    lines.extend(_render_oe(oe))
        lines.append("")

    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def convert_tree(xml_root: Path, md_root: Path) -> tuple[int, int]:
    """Walk ``xml_root`` for ``*.xml`` and write a mirror tree of ``.md``."""
    n_ok = 0
    n_err = 0
    for xml_path in xml_root.rglob("*.xml"):
        if xml_path.name.startswith("_"):
            continue
        rel = xml_path.relative_to(xml_root)
        md_path = (md_root / rel).with_suffix(".md")
        md_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            md_path.write_text(
                render_page(xml_path.read_text(encoding="utf-8")),
                encoding="utf-8",
            )
            n_ok += 1
        except Exception as exc:  # noqa: BLE001
            n_err += 1
            print(f"ERR {rel}: {exc}", file=sys.stderr)
    return n_ok, n_err
