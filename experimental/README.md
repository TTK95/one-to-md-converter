# Experimental: Markdown → OneNote writeback

This folder is for the inverse direction of the main tool: take a
Markdown file produced by `one-to-md export` and push the changes back
into OneNote via `Application.UpdatePageContent`.

## Status

Work in progress. `importer.py` is a Phase 1 sketch:

- It resolves the target OneNote page by matching the 8-character
  `md5(page_id)` suffix that `one_to_md.dumper` writes into filenames.
- It depends on a `md_to_xml` module (Markdown → OneNote 2013 XML) that
  is not yet bundled here. Until that lands, `importer.py` will fail at
  import.
- `UpdatePageContent` merges by `objectID`. The current sketch sends a
  *new* outline with no `objectID`s, so OneNote **adds** content rather
  than replacing it. A merge-aware Phase 2 needs to round-trip the
  original `objectID`s through the Markdown so it can identify which
  outline elements changed.

## Don't run this against production pages

For testing, point it at an empty page in a sandbox section
(e.g. `_Test Restructure`). Treat the importer as a code sketch until a
deterministic Phase 2 exists.
