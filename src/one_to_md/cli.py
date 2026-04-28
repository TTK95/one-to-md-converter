"""Command-line entry point for ``one-to-md-converter``."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from . import __version__
from .converter import convert_tree
from .dumper import dump_notebook


def _cmd_dump(args: argparse.Namespace) -> int:
    pages, errs = dump_notebook(args.notebook, Path(args.output))
    return 0 if errs == 0 else 1


def _cmd_convert(args: argparse.Namespace) -> int:
    src = Path(args.input)
    dst = Path(args.output)
    if not src.is_dir():
        print(f"Not a directory: {src}", file=sys.stderr)
        return 2
    n_ok, n_err = convert_tree(src, dst, emit_oids=not args.no_oids)
    print(f"Converted {n_ok} pages ({n_err} errors). Out: {dst}")
    return 0 if n_err == 0 else 1


def _cmd_export(args: argparse.Namespace) -> int:
    """Dump + convert in one shot, into a timestamped directory."""
    base = Path(args.output)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xml_dir = base / f"xml-{stamp}"
    md_dir = base / f"md-{stamp}"

    print(f"==> Dumping to {xml_dir}")
    pages, errs = dump_notebook(args.notebook, xml_dir)

    print(f"==> Converting to {md_dir}")
    n_ok, n_err = convert_tree(xml_dir, md_dir, emit_oids=not args.no_oids)
    print(f"Converted {n_ok} pages ({n_err} errors).")

    print()
    print(f"XML:      {xml_dir}")
    print(f"Markdown: {md_dir}")
    return 0 if errs == 0 and n_err == 0 else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="one-to-md",
        description=(
            "Convert OneNote desktop notebooks to Markdown. Requires OneNote "
            "to be running with the target notebook open."
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_dump = sub.add_parser("dump", help="Dump notebook pages to OneNote 2013 XML.")
    p_dump.add_argument("notebook", help="Notebook display name (as shown in OneNote).")
    p_dump.add_argument("output", help="Output directory for XML files.")
    p_dump.set_defaults(func=_cmd_dump)

    p_conv = sub.add_parser("convert", help="Convert a tree of XML pages to Markdown.")
    p_conv.add_argument("input", help="Directory holding the XML dump.")
    p_conv.add_argument("output", help="Output directory for Markdown files.")
    p_conv.add_argument(
        "--no-oids",
        action="store_true",
        help="Omit <!-- oid=... --> annotations (cleaner MD; not round-trippable).",
    )
    p_conv.set_defaults(func=_cmd_convert)

    p_exp = sub.add_parser("export", help="Dump + convert in one go.")
    p_exp.add_argument("notebook", help="Notebook display name.")
    p_exp.add_argument(
        "-o", "--output", default=".", help="Base directory for output (default: cwd)."
    )
    p_exp.add_argument(
        "--no-oids",
        action="store_true",
        help="Omit <!-- oid=... --> annotations (cleaner MD; not round-trippable).",
    )
    p_exp.set_defaults(func=_cmd_export)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
