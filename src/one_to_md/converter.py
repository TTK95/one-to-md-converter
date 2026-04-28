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


def _oid_comment(elem: ET.Element, indent: int = 0) -> str:
    oid = elem.attrib.get("objectID")
    if not oid:
        return ""
    return f"{'  ' * indent}<!-- oid={oid} -->"


def _render_oe(oe: ET.Element, indent: int = 0, emit_oids: bool = True) -> list[str]:
    """Render a OneNote outline element (and children) to Markdown lines.

    All OE oids ride inline as ``<span data-oid="…"></span>``. Block-level
    HTML comments would split any enclosing list at every annotated item,
    so we keep block-level annotations only for outline-level markers (handled
    in :func:`render_page`).
    """
    out: list[str] = []
    is_bullet = oe.find("on:List/on:Bullet", _NS) is not None
    is_numbered = oe.find("on:List/on:Number", _NS) is not None
    is_list = is_bullet or is_numbered
    pad = "   " * indent
    oid = oe.attrib.get("objectID") if emit_oids else None
    inline_oid = f'<span data-oid="{oid}"></span>' if oid else ""

    text_parts = [_render_t(t.text) for t in oe.findall("on:T", _NS) if t.text]
    text = " ".join(p for p in text_parts if p).strip()

    if text:
        if is_bullet:
            out.append(f"{pad}- {inline_oid}{text}")
        elif is_numbered:
            out.append(f"{pad}1. {inline_oid}{text}")
        else:
            out.append(f"{pad}{inline_oid}{text}")
        out.append("")
    else:
        img = oe.find("on:Image", _NS)
        ifile = oe.find("on:InsertedFile", _NS)
        if img is not None:
            placeholder = f"![{img.attrib.get('alt', 'image')}](image)"
            line = f"{pad}- {inline_oid}{placeholder}" if is_bullet else (
                f"{pad}1. {inline_oid}{placeholder}" if is_numbered else f"{pad}{inline_oid}{placeholder}"
            )
            out.append(line)
            out.append("")
        elif ifile is not None:
            placeholder = f"_(attached file: {ifile.attrib.get('preferredName', 'file')})_"
            line = f"{pad}- {inline_oid}{placeholder}" if is_bullet else (
                f"{pad}1. {inline_oid}{placeholder}" if is_numbered else f"{pad}{inline_oid}{placeholder}"
            )
            out.append(line)
            out.append("")
        elif inline_oid:
            if is_bullet:
                out.append(f"{pad}- {inline_oid}")
            elif is_numbered:
                out.append(f"{pad}1. {inline_oid}")
            else:
                out.append(f"{pad}{inline_oid}")
            out.append("")

    children = oe.find("on:OEChildren", _NS)
    if children is not None:
        nested_indent = indent + (1 if is_list else 0)
        for child_oe in children.findall("on:OE", _NS):
            out.extend(_render_oe(child_oe, nested_indent, emit_oids=emit_oids))

    return out


def _render_table(tbl: ET.Element, emit_oids: bool = True) -> list[str]:
    rows: list[list[str]] = []
    for row in tbl.findall("on:Row", _NS):
        cells: list[str] = []
        for cell in row.findall("on:Cell", _NS):
            cell_lines: list[str] = []
            for oechildren in cell.findall("on:OEChildren", _NS):
                for oe in oechildren.findall("on:OE", _NS):
                    cell_lines.extend(_render_oe(oe, emit_oids=False))
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


def render_page(xml_text: str, emit_oids: bool = True) -> str:
    """Convert a single OneNote 2013 ``<one:Page>`` document to Markdown.

    When ``emit_oids`` is true (default), each outline element is preceded by
    an ``<!-- oid=... -->`` HTML comment carrying its OneNote ``objectID``.
    The companion ``experimental.md_to_xml`` reads these comments back, so
    ``UpdatePageContent`` can merge edits by ID instead of appending.
    """
    root = ET.fromstring(xml_text)
    name = root.attrib.get("name", "Untitled")

    title_oid = ""
    title_elem = root.find("on:Title", _NS)
    if emit_oids and title_elem is not None:
        title_oe = title_elem.find("on:OE", _NS)
        if title_oe is not None and title_oe.attrib.get("objectID"):
            title_oid = f"<!-- oid={title_oe.attrib['objectID']} title=1 -->"

    lines: list[str] = []
    if title_oid:
        lines.append(title_oid)
    lines.extend([f"# {name}", ""])

    for child in root:
        tag = child.tag.split("}", 1)[-1]
        if tag != "Outline":
            continue
        if emit_oids and child.attrib.get("objectID"):
            lines.append(f"<!-- oid={child.attrib['objectID']} outline=1 -->")
        for oechildren in child.findall("on:OEChildren", _NS):
            for oe in oechildren.findall("on:OE", _NS):
                table = oe.find("on:Table", _NS)
                if table is not None:
                    if emit_oids and oe.attrib.get("objectID"):
                        rows = len(table.findall("on:Row", _NS))
                        cols = max(
                            (len(r.findall("on:Cell", _NS)) for r in table.findall("on:Row", _NS)),
                            default=0,
                        )
                        lines.append(
                            f"<!-- oid={oe.attrib['objectID']} table={cols}x{rows} -->"
                        )
                    lines.extend(_render_table(table, emit_oids=emit_oids))
                    lines.append("")
                else:
                    lines.extend(_render_oe(oe, emit_oids=emit_oids))
        lines.append("")

    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def convert_tree(xml_root: Path, md_root: Path, emit_oids: bool = True) -> tuple[int, int]:
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
                render_page(xml_path.read_text(encoding="utf-8"), emit_oids=emit_oids),
                encoding="utf-8",
            )
            n_ok += 1
        except Exception as exc:  # noqa: BLE001
            n_err += 1
            print(f"ERR {rel}: {exc}", file=sys.stderr)
    return n_ok, n_err
