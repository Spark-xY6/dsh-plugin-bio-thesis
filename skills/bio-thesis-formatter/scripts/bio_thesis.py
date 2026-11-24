#!/usr/bin/env python3
"""Unified CLI for thesis draft creation, audit and template margin repair."""

from __future__ import annotations

import argparse
import json
import pathlib
import sys

SCRIPTS_DIR = pathlib.Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import audit_thesis_docx  # noqa: E402
from docx_package import same_path, staged_file  # noqa: E402
import init_thesis_docx  # noqa: E402
import repair_margins  # noqa: E402

SKILL_DIR = SCRIPTS_DIR.parent.resolve()
CHECKLIST_DOCX = SKILL_DIR / "resources" / "bio-thesis-self-check-67.docx"
PRESETS_DIR = SKILL_DIR / "resources" / "presets"


def _cmd_export_checklist(args: argparse.Namespace) -> int:
    out_p = pathlib.Path(args.output).resolve()
    if out_p == SKILL_DIR or SKILL_DIR in out_p.parents:
        print(f"ERROR: Refusing to export checklist inside skill directory: {out_p}", file=sys.stderr)
        return 1
    if out_p.suffix.lower() != ".docx" or same_path(out_p, CHECKLIST_DOCX):
        print(f"ERROR: Output must be an independent .docx path: {out_p}", file=sys.stderr)
        return 1
    if out_p.exists() and not args.overwrite:
        print(f"ERROR: Output exists (pass --overwrite to replace): {out_p}", file=sys.stderr)
        return 1
    with staged_file(out_p, overwrite=args.overwrite) as stage:
        stage.write_bytes(CHECKLIST_DOCX.read_bytes())
    print(f"EXPORTED={out_p}")
    return 0


def _cmd_list_presets(_args: argparse.Namespace) -> int:
    rows = []
    for p in sorted(PRESETS_DIR.glob("*.json")):
        data = json.loads(p.read_text(encoding="utf-8"))
        rows.append(
            {
                "preset": data.get("preset_name", p.stem),
                "citation_style": data.get("citation_style"),
                "description": data.get("description"),
                "modes": list((data.get("modes") or {}).keys()),
                "path": str(p),
            }
        )
    print(json.dumps(rows, ensure_ascii=False, indent=2))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="bio_thesis.py",
        description="Institution-neutral Word (.docx) thesis formatter & 67-item compliance CLI (Python stdlib only).",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_init = sub.add_parser("init", help="Initialize a clean formatted thesis or report DOCX")
    init_thesis_docx.add_arguments(p_init)

    p_audit = sub.add_parser("audit", help="Run read-only OOXML & 67-item compliance audit on a DOCX")
    audit_thesis_docx.add_arguments(p_audit)

    margins = sub.add_parser("repair-margins", help="Preview four section margins against a school template; optionally save a copy")
    margins.add_argument("--docx", required=True, type=pathlib.Path, help="Existing thesis .docx")
    margins.add_argument("--template", required=True, type=pathlib.Path, help="Matching school .docx or .dotx")
    margins.add_argument("--output", type=pathlib.Path, help="New .docx file; omit for a read-only preview")

    full_repair = sub.add_parser(
        "repair",
        help="Align sections by role and preview/repair margins, header/footer distances, gutter, settings, page numbering, and optional tables/indents",
    )
    full_repair.add_argument("--docx", required=True, type=pathlib.Path, help="Existing thesis .docx")
    full_repair.add_argument("--template", required=True, type=pathlib.Path, help="School .docx or .dotx layout reference")
    full_repair.add_argument("--output", type=pathlib.Path, help="New .docx file; omit for a read-only preview")
    full_repair.add_argument("--no-fix-numbering", action="store_true", help="Keep existing pgNumType formats and restarts unchanged")
    full_repair.add_argument("--fix-tables", action="store_true", help="Normalize body tables to standard three-line borders and repeat header rows")
    full_repair.add_argument("--clean-leading-spaces", action="store_true", help="Strip stray leading spaces/tabs on indented body paragraphs")

    p_chk = sub.add_parser("export-checklist", help="Copy editable 67-item self-check DOCX to workspace")
    p_chk.add_argument("--output", required=True, help="Destination checklist .docx path")
    p_chk.add_argument("--overwrite", action="store_true", help="Overwrite existing destination")

    sub.add_parser("list-presets", help="List built-in formatting presets")

    args = parser.parse_args(argv)
    if args.command == "repair-margins":
        try:
            repair_margins.repair(args.docx, args.template, args.output)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
        return 0
    if args.command == "repair":
        try:
            repair_margins.repair_layout(
                args.docx,
                args.template,
                args.output,
                fix_numbering=not args.no_fix_numbering,
                fix_tables=args.fix_tables,
                clean_leading_spaces=args.clean_leading_spaces,
            )
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
        return 0
    if args.command == "init":
        try:
            init_thesis_docx.initialize_docx(args)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
        return 0
    if args.command == "audit":
        return audit_thesis_docx.run_parsed(args)
    if args.command == "export-checklist":
        return _cmd_export_checklist(args)
    if args.command == "list-presets":
        return _cmd_list_presets(args)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
