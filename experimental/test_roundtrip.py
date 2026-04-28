"""Round-trip sanity check: XML -> MD -> XML -> MD, expecting MD == MD.

Walks a directory of OneNote 2013 XML pages (as produced by ``one-to-md dump``),
converts each through ``one_to_md.converter.render_page`` to Markdown, then back
through ``experimental.md_to_xml.md_to_xml``, then re-renders to Markdown. If the
two Markdown texts disagree, the page is reported as a divergence.

This catches structural fidelity bugs in either direction without needing a live
OneNote instance.

Usage:
    python -m experimental.test_roundtrip <xml_dir> [--limit N] [--show-diff K]

Or, from inside experimental/:
    python test_roundtrip.py <xml_dir>
"""

from __future__ import annotations

import argparse
import difflib
import sys
from pathlib import Path

# Make sibling md_to_xml and parent src/ importable when run as a script.
_HERE = Path(__file__).resolve().parent
_REPO = _HERE.parent
sys.path.insert(0, str(_HERE))
sys.path.insert(0, str(_REPO / "src"))

from md_to_xml import md_to_xml  # noqa: E402
from one_to_md.converter import render_page  # noqa: E402


def _normalize(md: str) -> str:
    return "\n".join(line.rstrip() for line in md.splitlines() if line.strip()) + "\n"


def _page_id_from(xml_text: str) -> str:
    import re

    m = re.search(r'\bID="([^"]+)"', xml_text)
    return m.group(1) if m else "{00000000-0000-0000-0000-000000000000}{1}{B0}"


def _roundtrip_one(xml_path: Path) -> tuple[bool, str, str, str]:
    """Returns (ok, md1, md2, error_message)."""
    try:
        xml_text = xml_path.read_text(encoding="utf-8")
        md1 = render_page(xml_text, emit_oids=True)
        page_id = _page_id_from(xml_text)
        xml2 = md_to_xml(md1, page_id=page_id)
        md2 = render_page(xml2, emit_oids=True)
    except Exception as exc:
        return False, "", "", f"{type(exc).__name__}: {exc}"
    return _normalize(md1) == _normalize(md2), md1, md2, ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xml_dir", help="Directory of XML pages to round-trip.")
    ap.add_argument("--limit", type=int, default=0, help="Stop after N pages (0 = all).")
    ap.add_argument(
        "--show-diff",
        type=int,
        default=3,
        help="For the first K divergences, print a unified diff (default 3).",
    )
    args = ap.parse_args()

    root = Path(args.xml_dir)
    paths = [
        p
        for p in root.rglob("*.xml")
        if not p.name.startswith("_")
    ]
    if args.limit:
        paths = paths[: args.limit]

    n_ok = 0
    n_diff = 0
    n_err = 0
    diffs_shown = 0
    failing: list[tuple[Path, str]] = []

    for p in paths:
        ok, md1, md2, err = _roundtrip_one(p)
        if err:
            n_err += 1
            failing.append((p, err))
        elif ok:
            n_ok += 1
        else:
            n_diff += 1
            failing.append((p, "diverged"))
            if diffs_shown < args.show_diff:
                diffs_shown += 1
                print(f"\n=== DIFF {p.relative_to(root)} ===")
                diff = difflib.unified_diff(
                    _normalize(md1).splitlines(keepends=True),
                    _normalize(md2).splitlines(keepends=True),
                    fromfile="original",
                    tofile="roundtrip",
                    n=2,
                )
                sys.stdout.writelines(diff)

    print(
        f"\nTotal: {len(paths)}  ok={n_ok}  diverged={n_diff}  errors={n_err}"
    )
    if n_err:
        print("\nFirst 10 errors:")
        for p, e in [(p, e) for p, e in failing if e != "diverged"][:10]:
            print(f"  {p.relative_to(root)}: {e}")
    return 0 if (n_diff == 0 and n_err == 0) else 1


if __name__ == "__main__":
    sys.exit(main())
