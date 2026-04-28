# one-to-md-converter

Convert OneNote desktop notebooks to Markdown by talking to OneNote itself
through its COM API. Reliable on large notebooks (500+ pages tested) where
binary `.one` parsers fall over and notebook-level exporters crash OneNote.

## Why another OneNote → Markdown tool?

The pure binary-format parsers (`onenote.rs`, `one2html`, …) reject any
recent OneNote file with `Malformed FSSHTTPB data`. Tools that drive
OneNote via Word's "Publish" path (`OneNoteMdExporter`, …) crash OneNote
on bigger notebooks because they clone every page into a temporary
notebook first. And calling OneNote COM through `pywin32` /
`comtypes` returns *Library not registered* on Office click-to-run
installs that ship without the TypeLib registration.

This tool sidesteps all of that:

- It loads the in-GAC `Microsoft.Office.Interop.OneNote` assembly through
  [`pythonnet`](https://pypi.org/project/pythonnet/), which uses
  vtable-based .NET binding and does not need the TypeLib registry entry.
- It reads each page with `Application.GetPageContent` only — no temp
  notebooks, no Word round-trip — so OneNote stays responsive even for
  large notebooks.

## Requirements

- Windows.
- OneNote desktop installed (the classic Office app — *not* "OneNote for
  Windows 10"). Tested with OneNote 2016 / Microsoft 365.
- Python ≥ 3.9.
- The notebook you want to read **must already be open in OneNote**. The
  tool does not launch or sync OneNote for you.
- Run with normal user privileges, not as administrator. OneNote refuses
  to talk across an elevation boundary.

## Install

```bash
pip install --user .
```

This pulls in `pythonnet` and registers the `one-to-md` console script.

For a development checkout:

```bash
pip install --user -e .
```

## Usage

```bash
# 1) Open OneNote and make sure your notebook is loaded.

# 2) One-shot dump + convert into ./xml-<timestamp>/ and ./md-<timestamp>/
one-to-md export "My Notebook"

# Custom output base
one-to-md export "My Notebook" -o ./exports
```

Or run the two stages separately:

```bash
one-to-md dump "My Notebook" ./xml-out
one-to-md convert ./xml-out ./md-out
```

A run for a 507-page notebook completes in ~10 seconds.

## Output layout

```
md-out/
└── My Notebook/
    ├── Section A/
    │   ├── Page Title 1_<id8>.md
    │   └── …
    ├── Section Group X/
    │   └── Section B/
    │       └── …
    └── …
```

The 8-character suffix on each filename is the first 8 hex digits of
`md5(page_id)`. It keeps filenames stable across runs and uniquely
addresses pages even when titles collide.

## Conversion fidelity

Preserved: page titles, paragraphs, bullet/numbered lists with nesting,
tables, links, basic bold/italic, image and attachment placeholders.

Not preserved: ink/handwriting, embedded files (only a placeholder
note), audio, video, page positioning, fine-grained styling
(font/colour). The XML is also kept under `xml-…/` so you can extend the
converter to surface anything the default Markdown rendering drops.

## Limitations

- Only sees what OneNote desktop sees: notebooks must already be open
  and synced.
- Password-protected sections are skipped unless they are unlocked
  before the run.
- The free *OneNote for Windows 10* app is not supported (it does not
  expose the COM API).

## Experimental: Markdown → OneNote (writeback)

`experimental/` ships a working writeback path: edit the Markdown copy of a
page and push the changes back into OneNote via `Application.UpdatePageContent`.

```bash
one-to-md convert ./xml-out ./md-out         # MD now carries oid annotations
# ... edit ./md-out/.../My Page_abcd1234.md ...
python experimental/importer.py "My Notebook" ./md-out/.../My\ Page_abcd1234.md
```

By default the conversion preserves OneNote `objectID`s as `<!-- oid=… -->`
HTML comments and `<span data-oid="…"></span>` inline markers, so
`UpdatePageContent` *merges by ID* — edited blocks replace, untouched blocks
stay put. Pass `--no-oids` to `convert` for clean Markdown if you don't need
writeback. See [`experimental/README.md`](experimental/README.md) for schema
coverage, round-trip fidelity numbers, and live-test guidance.

## License

MIT — see [`LICENSE`](LICENSE).
