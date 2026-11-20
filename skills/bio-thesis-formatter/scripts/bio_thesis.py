#!/usr/bin/env python3
"""Unified zero-dependency CLI for bio-thesis-formatter (init, audit, export-checklist, list-presets)."""

from __future__ import annotations

import argparse
import json
import pathlib
import shutil
import sys

SCRIPTS_DIR = pathlib.Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import audit_thesis_docx  # noqa: E402
import init_thesis_docx  # noqa: E402

SKILL_DIR = SCRIPTS_DIR.parent
CHECKLIST_DOCX = SKILL_DIR / "resources" / "bio-thesis-self-check-65.docx"
PRESETS_DIR = SKILL_DIR / "resources" / "presets"


def _cmd_export_checklist(args: argparse.Namespace) -> int:
    out_p = pathlib.Path(args.output).resolve()
    if out_p.exists() and not args.overwrite:
        print(f"ERROR: Output exists (pass --overwrite to replace): {out_p}", file=sys.stderr)
        return 1
    out_p.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(CHECKLIST_DOCX, out_p)
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
        description="Institution-neutral Word (.docx) thesis formatter & 65-item compliance CLI (Python stdlib only).",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_init = sub.add_parser("init", help="Initialize a clean formatted thesis or report DOCX")
    init_thesis_docx.add_arguments(p_init)

    p_audit = sub.add_parser("audit", help="Run read-only OOXML & 65-item compliance audit on a DOCX")
    audit_thesis_docx.add_arguments(p_audit)

    p_chk = sub.add_parser("export-checklist", help="Copy editable 65-item self-check DOCX to workspace")
    p_chk.add_argument("--output", required=True, help="Destination checklist .docx path")
    p_chk.add_argument("--overwrite", action="store_true", help="Overwrite existing destination")

    sub.add_parser("list-presets", help="List built-in formatting presets")

    args = parser.parse_args(argv)
    if args.command == "init":
        try:
            init_thesis_docx.initialize_docx(args)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
        return 0
    if args.command == "audit":
        return audit_thesis_docx.main(
            [
                "--docx", args.docx,
                "--mode", args.mode,
                "--review-stage", args.review_stage,
                "--preset", args.preset,
                "--citation-style", args.citation_style,
                *(["--fail-on-warn"] if args.fail_on_warn else []),
                *(["--template", args.template] if args.template else []),
                *(["--base-docx", args.base_docx] if args.base_docx else []),
                *(["--profile-config", args.profile_config] if args.profile_config else []),
                *(["--genes", args.genes] if args.genes else []),
                *(["--species", args.species] if args.species else []),
                *(["--sensitive-names", args.sensitive_names] if args.sensitive_names else []),
                *(["--json-report", args.json_report] if args.json_report else []),
                *(["--md-report", args.md_report] if args.md_report else []),
            ]
        )
    if args.command == "export-checklist":
        return _cmd_export_checklist(args)
    if args.command == "list-presets":
        return _cmd_list_presets(args)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
