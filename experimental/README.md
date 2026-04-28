# Experimental: Markdown → OneNote writeback

The inverse of the main `one-to-md export` path. Take a Markdown file produced
by `one-to-md convert`, parse it back to OneNote 2013 XML, and push it onto the
matching page via `Application.UpdatePageContent`.

## Pieces

- `md_to_xml.py` — Markdown → OneNote 2013 XML page builder.
- `importer.py` — CLI driver that resolves the target page from the filename's
  8-char hash and calls `UpdatePageContent`.
- `test_roundtrip.py` — non-COM fidelity check: runs `XML → MD → XML → MD` and
  diffs the two MD outputs.

## Two modes

Whether the writeback *replaces* matching blocks or *appends* depends on
whether the markdown carries object-ID annotations.

### Phase 2 — merge-aware (recommended)

`one-to-md convert` (and `export`) now write `<!-- oid=… -->` and
`<span data-oid="…"></span>` annotations into the markdown by default. Each
title, outline, OE, and table block carries its OneNote `objectID`. Bullet and
numbered list items carry the oid as an inline span (HTML comments would split
the list at every annotated item). On writeback, OneNote merges by `objectID`:
edits replace the matching block, untouched blocks stay put.

This is the mode you want for editing real pages. If you delete an annotation
by hand the block degrades to additive — its content gets appended rather than
replacing — but no other block is affected.

To opt out and produce clean (non-annotated) Markdown, pass `--no-oids` to
`one-to-md convert` / `export`. That MD is for reading only; sending it back
through the importer triggers Phase 1 behavior.

### Phase 1 — additive

If the markdown has no oid annotations (e.g. the user wrote it from scratch,
or `--no-oids` was used), `UpdatePageContent` cannot match any blocks and
appends everything to the end of the page. Only safe against an empty sandbox
page.

## Usage

```bash
# 1) Make sure OneNote desktop is running with the notebook open.
# 2) Convert and edit:
one-to-md convert ./xml-out ./md-out
# ... edit the .md files in place ...

# 3) Push edits back:
python experimental/importer.py "My Notebook" ./md-out/My\ Notebook/Section/Page_abcd1234.md

# Inspect the generated XML without touching OneNote:
python experimental/importer.py "My Notebook" path/to/page.md --dry-run --debug-xml /tmp/page.xml
```

The filename's `_<8charhash>.md` suffix is the first 8 hex digits of
`md5(page_id)`. The importer reads it to look the target page up via
`GetHierarchy`. If two pages in one notebook collide on the suffix the
importer aborts with the candidates listed (mathematically rare for a single
notebook of normal size, but cheap to detect).

## Round-trip fidelity check

`test_roundtrip.py` walks an XML dump, runs each page through both directions,
and reports any pages whose Markdown differs after a round trip:

```bash
python experimental/test_roundtrip.py ./xml-out
python experimental/test_roundtrip.py ./xml-out --limit 20 --show-diff 3
```

On the reference 508-page notebook the test reports ~93% byte-identical round
trips. Remaining diffs are markdown-it lexer cosmetics:

- backslash-escaped Markdown specials in plain text (`\_`, `\\`) collapse on
  re-parse — content is preserved at the XML level, only the MD source differs;
- alphabetic ordered lists (`a. …`) inside list items aren't part of CommonMark,
  so indentation falls back to a paragraph;
- a few link URLs round-trip with normalized whitespace.

These don't affect writeback fidelity — the regenerated XML is functionally
equivalent.

## Schema coverage

| OneNote element | Round-tripped | Notes |
|---|---|---|
| Page title (`one:Title`) | ✓ | first H1 in MD; oid kept on title OE. |
| Outlines (`one:Outline`) | ✓ | each outline-level oid emitted as `<!-- oid=… outline=1 -->`; multiple outlines preserved. |
| Paragraphs | ✓ | inline oid via `<span data-oid="…"></span>`. |
| Bullet / numbered lists, nested | ✓ | indented 3 spaces per level (works for both `-` and `1.`); oids ride inline. |
| Tables | ✓ | block-level oid annotation includes `cols=N rows=M`. |
| Inline formatting (`**`, `*`, links, `~~`, code) | ✓ | inverse of `converter._render_t`. |
| Image / file placeholders | ✓ (with oid) | `![alt](image)` and `_(attached file: name)_`. With an oid present, an OE with `<one:Image>` / `<one:InsertedFile>` is sent so OneNote keeps the attachment. Without an oid the block is dropped — adding new attachments via writeback isn't supported. |
| Empty OEs | ✓ | preserved as `<one:OE objectID="X"><one:T><![CDATA[]]></one:T></one:OE>` so subsequent merges leave the original block intact. |
| Tags / todo checkboxes (`one:Tag`) | ✗ | dropped on the way out, not restored. |
| Ink, audio, video, MediaIndex | ✗ | dropped on the way out. |
| Page settings (color, rule lines) | ✗ | OneNote uses defaults on writeback if the original is replaced. |

## Live writeback testing

For the first run against a real page, work in a sandbox section
(e.g. `_Test Restructure`):

1. Create an empty page in the sandbox section.
2. `one-to-md export "<Notebook>" -o ./out`.
3. Edit one block in the Markdown copy of that page (e.g. change a bullet's
   text). Keep all `<!-- oid=… -->` and `<span data-oid="…">` annotations.
4. `python experimental/importer.py "<Notebook>" path/to/sandbox/page.md --debug-xml /tmp/check.xml`.
5. Re-export and diff against the previous dump. The edited block should have
   changed, every other block should be byte-identical.

This is the fastest way to confirm `UpdatePageContent` is honoring the
`objectID`s the way we expect on your specific OneNote install. Schema
rejections (rare on the 2013 schema) surface as a `COMException` from
`UpdatePageContent`; the saved `--debug-xml` is your starting point.
