"""Markdown -> OneNote 2013 XML page builder.

Inverse of :mod:`one_to_md.converter`. Used by ``experimental/importer.py`` to
push edited Markdown back into a OneNote page via
``Application.UpdatePageContent``.

Phase 1 (additive): emits a complete ``<one:Page>`` document but without
``objectID``s on outline elements -- ``UpdatePageContent`` will *append* the
content rather than replace matching blocks. Use against an empty sandbox
page.

Phase 2 (merge-aware): when the input markdown has ``<!-- oid={...}{n}{B0} -->``
comments before each block (as emitted by ``converter.render_page`` with
``emit_oids=True``), those object IDs are placed on the corresponding outline
elements, allowing ``UpdatePageContent`` to merge by ``objectID``.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional

from markdown_it import MarkdownIt
from markdown_it.token import Token


NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

_OID_RE = re.compile(r"<!--\s*oid=([^\s>]+)(?:\s+([^>]+?))?\s*-->")
_INLINE_OID_RE = re.compile(r'<span\s+data-oid="([^"]+)"\s*/?>')
_PLACEHOLDER_IMAGE_RE = re.compile(r"^!\[([^\]]*)\]\(image\)\s*$")
_PLACEHOLDER_FILE_RE = re.compile(r"^_\(attached file: ([^)]+)\)_\s*$")


@dataclass
class _Block:
    """Internal IR for a OneNote outline element before serialization."""

    kind: str  # 'title' | 'h2'..'h6' | 'p' | 'bullet' | 'number' | 'table' | 'placeholder' | 'empty'
    text_html: str = ""
    children: List["_Block"] = field(default_factory=list)
    oid: Optional[str] = None
    extra: dict = field(default_factory=dict)


@dataclass
class _ParseState:
    """Accumulator threaded through token walking."""

    pending_oid: Optional[str] = None
    pending_extra: dict = field(default_factory=dict)


def _esc(s: str) -> str:
    return html.escape(s, quote=False)


def _esc_attr(s: str) -> str:
    return html.escape(s, quote=True)


def _extract_inline_oid(children: List[Token]) -> tuple[Optional[str], List[Token]]:
    """Strip leading ``<span data-oid="…"></span>`` markers from inline children.

    Returns the extracted oid (or ``None``) and the remaining children. A
    matched opening tag also consumes the matching ``</span>`` close so the
    caller never sees a dangling close in the rendered text.
    """
    oid: Optional[str] = None
    out: List[Token] = []
    skip_next_close = False
    for t in children:
        if t.type == "html_inline":
            content = t.content.strip()
            if skip_next_close and content == "</span>":
                skip_next_close = False
                continue
            m = _INLINE_OID_RE.fullmatch(content)
            if m and oid is None:
                oid = m.group(1)
                if not content.endswith("/>"):
                    skip_next_close = True
                continue
        out.append(t)
    return oid, out


def _render_inline(children: Iterable[Token]) -> str:
    """Render markdown-it inline children to a CDATA-safe HTML fragment.

    Inverts :func:`one_to_md.converter._render_t`.
    """
    out: List[str] = []
    for t in children:
        ty = t.type
        if ty == "text":
            out.append(_esc(t.content))
        elif ty == "strong_open":
            out.append("<span style='font-weight:bold'>")
        elif ty == "strong_close":
            out.append("</span>")
        elif ty == "em_open":
            out.append("<span style='font-style:italic'>")
        elif ty == "em_close":
            out.append("</span>")
        elif ty == "s_open":
            out.append("<span style='text-decoration:line-through'>")
        elif ty == "s_close":
            out.append("</span>")
        elif ty == "code_inline":
            out.append("<span style='font-family:Consolas'>")
            out.append(_esc(t.content))
            out.append("</span>")
        elif ty == "link_open":
            href = t.attrGet("href") or ""
            out.append(f'<a href="{_esc_attr(href)}">')
        elif ty == "link_close":
            out.append("</a>")
        elif ty == "softbreak":
            out.append("\n")
        elif ty == "hardbreak":
            out.append("<br/>")
        elif ty == "image":
            alt = t.content or t.attrGet("alt") or ""
            src = t.attrGet("src") or ""
            out.append(f"![{_esc(alt)}]({_esc(src)})")
        elif ty == "html_inline":
            out.append(t.content)
        else:
            if t.content:
                out.append(_esc(t.content))
    return "".join(out)


def _consume_oid_comment(content: str) -> tuple[Optional[str], dict]:
    m = _OID_RE.search(content)
    if not m:
        return None, {}
    oid = m.group(1)
    extra: dict = {}
    if m.group(2):
        for part in m.group(2).split():
            if "=" in part:
                k, v = part.split("=", 1)
                extra[k] = v
    return oid, extra


def _slice_balanced(
    tokens: List[Token], start: int, open_type: str, close_type: str
) -> int:
    """Return index of the matching close token at depth 0."""
    depth = 1
    i = start
    while i < len(tokens):
        ty = tokens[i].type
        if ty == open_type:
            depth += 1
        elif ty == close_type:
            depth -= 1
            if depth == 0:
                return i
        i += 1
    raise ValueError(f"Unbalanced {open_type}/{close_type} starting at {start}")


def _parse_table(tokens: List[Token], start: int) -> tuple[_Block, int]:
    """Parse a table_open ... table_close span. Returns (block, index_after_close)."""
    end = _slice_balanced(tokens, start + 1, "table_open", "table_close")
    rows: List[List[str]] = []
    current_row: List[str] = []
    i = start + 1
    while i < end:
        ty = tokens[i].type
        if ty == "tr_open":
            current_row = []
        elif ty == "tr_close":
            rows.append(current_row)
        elif ty in ("th_open", "td_open"):
            inline = tokens[i + 1]
            current_row.append(_render_inline(inline.children or []))
            i += 2  # skip inline + close
        i += 1
    block = _Block(kind="table", extra={"rows": rows})
    return block, end + 1


def _flush_pending(state: _ParseState, blocks: List[_Block]) -> None:
    """Emit an empty-OE placeholder for a pending oid that nothing consumed."""
    if state.pending_oid is not None:
        blocks.append(
            _Block(
                kind="empty", oid=state.pending_oid, extra=dict(state.pending_extra)
            )
        )
        state.pending_oid = None
        state.pending_extra = {}


def _absorb_oid_comment(state: _ParseState, blocks: List[_Block], content: str) -> None:
    """Update parser state from an html_block carrying an ``oid=`` comment.

    Outline-level oids become ``outline_marker`` blocks so the serializer can
    split body content into multiple ``<one:Outline>`` elements. Other oids
    overwrite ``pending_oid`` after flushing any prior unconsumed pending.
    """
    oid, extra = _consume_oid_comment(content)
    if oid is None:
        return
    if extra.get("outline"):
        _flush_pending(state, blocks)
        blocks.append(_Block(kind="outline_marker", oid=oid))
        return
    _flush_pending(state, blocks)
    state.pending_oid = oid
    state.pending_extra = extra


def _parse_list(
    tokens: List[Token], start: int, parent_state: Optional[_ParseState] = None
) -> tuple[List[_Block], int]:
    """Parse a ``*_list_open ... *_list_close`` span into list-item blocks.

    The first item inherits any unconsumed ``pending_oid`` from the parent
    walker so that an oid comment immediately preceding the list lands on the
    first item rather than being lost when control passes into this function.
    """
    open_t = tokens[start]
    open_type = open_t.type
    close_type = open_type.replace("_open", "_close")
    kind = "bullet" if open_type == "bullet_list_open" else "number"
    end = _slice_balanced(tokens, start + 1, open_type, close_type)

    items: List[_Block] = []
    state = _ParseState()
    if parent_state is not None:
        state.pending_oid = parent_state.pending_oid
        state.pending_extra = dict(parent_state.pending_extra)
        parent_state.pending_oid = None
        parent_state.pending_extra = {}
    i = start + 1
    while i < end:
        ty = tokens[i].type
        if ty == "html_block":
            _absorb_oid_comment(state, items, tokens[i].content)
            i += 1
            continue
        if ty == "list_item_open":
            li_end = _slice_balanced(tokens, i + 1, "list_item_open", "list_item_close")
            inner_blocks = _parse_blocks(tokens[i + 1 : li_end])
            head_text = ""
            head_oid: Optional[str] = None
            children: List[_Block] = []
            if inner_blocks:
                first = inner_blocks[0]
                if first.kind == "p":
                    head_text = first.text_html
                    head_oid = first.oid
                    children = inner_blocks[1:]
                else:
                    children = inner_blocks
            item_oid = head_oid or state.pending_oid
            items.append(
                _Block(
                    kind=kind,
                    text_html=head_text,
                    children=children,
                    oid=item_oid,
                    extra=dict(state.pending_extra) if not head_oid else {},
                )
            )
            state.pending_oid = None
            state.pending_extra = {}
            i = li_end + 1
            continue
        i += 1
    _flush_pending(state, items)
    return items, end + 1


def _parse_blocks(
    tokens: List[Token], state: Optional[_ParseState] = None
) -> List[_Block]:
    blocks: List[_Block] = []
    if state is None:
        state = _ParseState()
    i = 0

    while i < len(tokens):
        t = tokens[i]
        ty = t.type

        if ty == "html_block":
            _absorb_oid_comment(state, blocks, t.content)
            i += 1
            continue

        if ty == "heading_open":
            level = int(t.tag[1])
            inline = tokens[i + 1]
            text = _render_inline(inline.children or [])
            kind = "title" if level == 1 else f"h{min(level, 6)}"
            blocks.append(
                _Block(
                    kind=kind,
                    text_html=text,
                    oid=state.pending_oid,
                    extra=dict(state.pending_extra),
                )
            )
            state.pending_oid = None
            state.pending_extra = {}
            i += 3
            continue

        if ty == "paragraph_open":
            inline = tokens[i + 1]
            inline_oid, inline_children = _extract_inline_oid(inline.children or [])
            text = _render_inline(inline_children)
            kind = "p"
            stripped = re.sub(r"<[^>]+>", "", text).strip()
            if _PLACEHOLDER_IMAGE_RE.match(stripped) or _PLACEHOLDER_FILE_RE.match(
                stripped
            ):
                kind = "placeholder"
            block_oid = inline_oid or state.pending_oid
            blocks.append(
                _Block(
                    kind=kind,
                    text_html=text,
                    oid=block_oid,
                    extra=dict(state.pending_extra) if not inline_oid else {},
                )
            )
            state.pending_oid = None
            state.pending_extra = {}
            i += 3
            continue

        if ty in ("bullet_list_open", "ordered_list_open"):
            items, new_i = _parse_list(tokens, i, state)
            blocks.extend(items)
            i = new_i
            continue

        if ty == "table_open":
            tbl, new_i = _parse_table(tokens, i)
            tbl.oid = state.pending_oid
            tbl.extra.update(state.pending_extra)
            state.pending_oid = None
            state.pending_extra = {}
            blocks.append(tbl)
            i = new_i
            continue

        if ty in ("fence", "code_block"):
            content = t.content.rstrip("\n")
            html_text = (
                "<span style='font-family:Consolas'>"
                + _esc(content).replace("\n", "<br/>")
                + "</span>"
            )
            blocks.append(
                _Block(
                    kind="p",
                    text_html=html_text,
                    oid=state.pending_oid,
                    extra=dict(state.pending_extra),
                )
            )
            state.pending_oid = None
            state.pending_extra = {}
            i += 1
            continue

        if ty == "hr":
            blocks.append(
                _Block(
                    kind="p",
                    text_html="<hr/>",
                    oid=state.pending_oid,
                    extra=dict(state.pending_extra),
                )
            )
            state.pending_oid = None
            state.pending_extra = {}
            i += 1
            continue

        if ty == "blockquote_open":
            bq_end = _slice_balanced(tokens, i + 1, "blockquote_open", "blockquote_close")
            inner = _parse_blocks(tokens[i + 1 : bq_end])
            for ib in inner:
                if ib.kind == "p":
                    ib.text_html = (
                        "<span style='font-style:italic'>" + ib.text_html + "</span>"
                    )
                blocks.append(ib)
            i = bq_end + 1
            continue

        i += 1

    _flush_pending(state, blocks)
    return blocks


def _serialize_oechildren(blocks: List[_Block]) -> str:
    if not blocks:
        return ""
    inner = "".join(_serialize_block(b) for b in blocks)
    if not inner:
        return ""
    return f"<one:OEChildren>{inner}</one:OEChildren>"


def _serialize_block(b: _Block) -> str:
    oid_attr = f' objectID="{_esc_attr(b.oid)}"' if b.oid else ""

    if b.kind == "empty":
        if not b.oid:
            return ""
        return f"<one:OE{oid_attr}><one:T><![CDATA[]]></one:T></one:OE>"

    if b.kind == "placeholder":
        if not b.oid:
            return ""
        stripped = re.sub(r"<[^>]+>", "", b.text_html).strip()
        m_img = _PLACEHOLDER_IMAGE_RE.match(stripped)
        m_file = _PLACEHOLDER_FILE_RE.match(stripped)
        if m_img is not None:
            alt = html.unescape(m_img.group(1))
            return (
                f"<one:OE{oid_attr}>"
                f'<one:Image alt="{_esc_attr(alt)}"/>'
                f"</one:OE>"
            )
        if m_file is not None:
            name = html.unescape(m_file.group(1))
            return (
                f"<one:OE{oid_attr}>"
                f'<one:InsertedFile preferredName="{_esc_attr(name)}"/>'
                f"</one:OE>"
            )
        return f"<one:OE{oid_attr}><one:T><![CDATA[]]></one:T></one:OE>"

    if b.kind == "title":
        return (
            f"<one:Title>"
            f'<one:OE{oid_attr} quickStyleIndex="0">'
            f"<one:T><![CDATA[{b.text_html}]]></one:T>"
            f"</one:OE>"
            f"</one:Title>"
        )

    if b.kind in ("h2", "h3", "h4", "h5", "h6"):
        idx = int(b.kind[1])
        return (
            f'<one:OE{oid_attr} quickStyleIndex="{idx}">'
            f"<one:T><![CDATA[{b.text_html}]]></one:T>"
            f"</one:OE>"
        )

    if b.kind == "p":
        children_xml = _serialize_oechildren(b.children)
        return (
            f"<one:OE{oid_attr}>"
            f"<one:T><![CDATA[{b.text_html}]]></one:T>"
            f"{children_xml}"
            f"</one:OE>"
        )

    if b.kind in ("bullet", "number"):
        marker = "<one:Bullet/>" if b.kind == "bullet" else '<one:Number/>'
        children_xml = _serialize_oechildren(b.children)
        return (
            f"<one:OE{oid_attr}>"
            f"<one:List>{marker}</one:List>"
            f"<one:T><![CDATA[{b.text_html}]]></one:T>"
            f"{children_xml}"
            f"</one:OE>"
        )

    if b.kind == "table":
        rows = b.extra.get("rows", [])
        row_xml: List[str] = []
        for r in rows:
            cells_xml: List[str] = []
            for cell_html in r:
                cells_xml.append(
                    f"<one:Cell>"
                    f"<one:OEChildren>"
                    f"<one:OE><one:T><![CDATA[{cell_html}]]></one:T></one:OE>"
                    f"</one:OEChildren>"
                    f"</one:Cell>"
                )
            row_xml.append(f"<one:Row>{''.join(cells_xml)}</one:Row>")
        return (
            f"<one:OE{oid_attr}>"
            f"<one:Table>{''.join(row_xml)}</one:Table>"
            f"</one:OE>"
        )

    return ""


def _make_md() -> MarkdownIt:
    md = MarkdownIt("commonmark").enable("table").enable("strikethrough")
    # Don't percent-encode link URLs -- OneNote links carry literal `{`, `\`,
    # and unicode characters that we want to preserve verbatim for round-trip
    # fidelity.
    md.normalizeLink = lambda url: url
    md.normalizeLinkText = lambda url: url
    md.validateLink = lambda url: True
    return md


def md_to_xml(
    md_text: str,
    page_id: str,
    *,
    page_name: Optional[str] = None,
    page_lang: str = "de",
) -> str:
    """Convert Markdown to a OneNote 2013 ``<one:Page>`` XML document.

    Args:
        md_text: Markdown source.
        page_id: ``ID`` attribute for the resulting ``<one:Page>``. The importer
            resolves this from the filename's 8-char hash suffix.
        page_name: Override the page title. If ``None`` (default), use the first
            H1 from the markdown, falling back to ``"Untitled"``.
        page_lang: ``lang`` attribute on the page.

    Returns the XML document as a string.
    """
    tokens = _make_md().parse(md_text)
    state = _ParseState()
    blocks = _parse_blocks(tokens, state)

    title_block: Optional[_Block] = None
    body_blocks: List[_Block] = []
    for b in blocks:
        if b.kind == "title" and title_block is None:
            title_block = b
        elif b.kind == "title":
            b.kind = "h2"
            body_blocks.append(b)
        else:
            body_blocks.append(b)

    if page_name is None:
        if title_block is not None:
            stripped = re.sub(r"<[^>]+>", "", title_block.text_html).strip()
            page_name = html.unescape(stripped) or "Untitled"
        else:
            page_name = "Untitled"

    title_xml = _serialize_block(title_block) if title_block else ""

    # Group body blocks into outline segments delimited by outline_marker blocks.
    segments: List[tuple[Optional[str], List[_Block]]] = []
    current_oid: Optional[str] = None
    current_blocks: List[_Block] = []

    def _flush_segment() -> None:
        if current_oid is not None or current_blocks:
            segments.append((current_oid, current_blocks))

    for b in body_blocks:
        if b.kind == "outline_marker":
            _flush_segment()
            current_oid = b.oid
            current_blocks = []
        else:
            current_blocks.append(b)
    _flush_segment()

    body_xml_parts: List[str] = []
    for oid, segs in segments:
        inner = "".join(_serialize_block(b) for b in segs)
        if not inner and not oid:
            continue
        attr = f' objectID="{_esc_attr(oid)}"' if oid else ""
        if not inner:
            inner = "<one:OEChildren/>"
        else:
            inner = f"<one:OEChildren>{inner}</one:OEChildren>"
        body_xml_parts.append(f"<one:Outline{attr}>{inner}</one:Outline>")
    body_xml = "".join(body_xml_parts)

    page_attrs = (
        f'xmlns:one="{NS}" '
        f'ID="{_esc_attr(page_id)}" '
        f'name="{_esc_attr(page_name)}" '
        f'lang="{_esc_attr(page_lang)}"'
    )

    return (
        '<?xml version="1.0"?>'
        f"<one:Page {page_attrs}>{title_xml}{body_xml}</one:Page>"
    )


__all__ = ["md_to_xml", "NS"]
