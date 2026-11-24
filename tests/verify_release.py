#!/usr/bin/env python3
"""De-identification, clean-room OOXML, multi-mode E2E, and hygiene release gate."""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import zlib
import struct

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent.resolve()
SKILL_DIR = REPO_ROOT / "skills" / "bio-thesis-formatter"
CLI_SCRIPT = SKILL_DIR / "scripts" / "bio_thesis.py"
MASTER_DOCX = SKILL_DIR / "resources" / "bio-thesis-master.docx"
CHECKLIST_DOCX = SKILL_DIR / "resources" / "bio-thesis-self-check-67.docx"
CUSTOM_PRESET = REPO_ROOT / "presets" / "custom-university.example.json"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

TEXT_EXTENSIONS = {
    "", ".css", ".html", ".js", ".json", ".md", ".mjs", ".ps1", ".py",
    ".sh", ".svg", ".toml", ".ts", ".txt", ".xml", ".yaml", ".yml",
}
USER_HOME_PATH = re.compile(
    r"(?:[a-z]:[\\/]+users[\\/]+|(?<![\w:/])/(?:home|Users)/)[\w.-]+",
    re.IGNORECASE,
)


def release_paths():
    """Yield first-party paths, pruning VCS metadata and installed dependencies."""
    for directory, dirs, names in os.walk(REPO_ROOT):
        dirs[:] = sorted(d for d in dirs if d not in {".git", "node_modules"})
        for name in [*dirs, *sorted(n for n in names if n != ".git")]:
            path = pathlib.Path(directory, name)
            if path.is_symlink():
                raise ValueError("Unexpected symlink in the first-party release tree")
            yield path


def load_private_terms(path: pathlib.Path | None) -> list[str]:
    if path is None:
        return []
    if path.resolve().is_relative_to(REPO_ROOT):
        raise ValueError("The private terms file must be outside the repository")
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    return sorted({line.strip().casefold() for line in lines if line.strip()})


def _make_1x1_png(out_png: pathlib.Path) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00\x0f\x4c\x81")
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")
    out_png.write_bytes(png)


def _run_cli(args: list[str], label: str) -> str:
    proc = subprocess.run(
        [sys.executable, "-B", "-X", "utf8", str(CLI_SCRIPT), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    print(f"\n=== {label} (rc={proc.returncode}) ===")
    if proc.stdout.strip():
        print(proc.stdout.strip())
    if proc.stderr.strip():
        print("STDERR:", proc.stderr.strip())
    assert proc.returncode == 0, f"{label} failed with rc={proc.returncode}"
    return proc.stdout


def check_deidentification_and_ns(terms: list[str]) -> None:
    def scan(text: str, label: str) -> None:
        folded = text.casefold()
        if USER_HOME_PATH.search(text) or any(term in folded for term in terms):
            raise ValueError(f"Private term or local user-home path found in {label}")

    for p in release_paths():
        rel = p.relative_to(REPO_ROOT).as_posix()
        scan(rel, "a repository path")
        if not p.is_file():
            continue
        if p.suffix.lower() in {".docx", ".dotx"}:
            with zipfile.ZipFile(p, "r") as zf:
                assert zf.testzip() is None, f"Corrupt ZIP: {rel}"
                media = [m for m in zf.namelist() if m.startswith("word/media/")]
                if p in {MASTER_DOCX, CHECKLIST_DOCX}:
                    assert not media, f"Shipped neutral resource {rel} must have 0 media images"
                for member in zf.namelist():
                    scan(member, f"an archive member name in {rel}")
                    if member.endswith((".xml", ".rels")):
                        txt = zf.read(member).decode("utf-8", errors="replace")
                        assert "ns0:" not in txt, f"Unexpected ns0: prefix in {rel}:{member}"
                        scan(txt, f"{rel}:{member}")
        elif p.suffix.lower() in TEXT_EXTENSIONS:
            scan(p.read_text(encoding="utf-8", errors="replace"), rel)
    print(f"GATE-1 RELEASE PRIVACY & OOXML: PASS (user-home path patterns, "
          f"{len(terms)} supplied private terms, neutral resources without media)")


def check_checklist() -> None:
    with zipfile.ZipFile(CHECKLIST_DOCX, "r") as zf:
        assert zf.testzip() is None
        raw = zf.read("word/document.xml")
        assert b"EN.CITE" not in raw
        root = ET.fromstring(raw)
    tbl = root.find(f".//{W_NS and f'{{{W_NS}}}tbl'}")
    assert tbl is not None
    sections: dict[str, int] = {}
    cur = None
    for tr in tbl.findall(f"./{{{W_NS}}}tr"):
        cells = ["".join(t.text or "" for t in tc.iter(f"{{{W_NS}}}t")).strip() for tc in tr.findall(f"./{{{W_NS}}}tc")]
        if len(cells) == 1 and cells[0]:
            cur = cells[0]
            sections[cur] = 0
        elif cur and cells and cells[0].isdigit():
            sections[cur] += 1
    assert sum(sections.values()) == 67, f"Expected 67 checklist items, got {sections}"
    print(f"GATE-2 CHECKLIST: PASS (sections={sections}, total=67)")


def check_repo_hygiene() -> None:
    bad = []
    count = 0
    for p in release_paths():
        if "__pycache__" in p.parts or p.suffix in (".pyc", ".pyo", ".tmp", ".bak"):
            bad.append(str(p))
        elif p.is_file():
            count += 1
    assert not bad, f"Hygiene check found disallowed files: {bad}"
    print(f"GATE-4 HYGIENE: PASS ({count} repo files, 0 cache/tmp artifacts)")


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify dsh-plugin-bio-thesis release readiness.")
    parser.add_argument("--out-dir", help="Optional scratch directory outside repo for E2E outputs")
    parser.add_argument("--private-terms-file", type=pathlib.Path,
                        help="Optional UTF-8 file outside the repository, one term per line")
    args = parser.parse_args()

    if not __debug__:
        parser.error("Release assertions require Python without optimization")
    if args.out_dir and pathlib.Path(args.out_dir).resolve().is_relative_to(REPO_ROOT):
        parser.error("--out-dir must be outside the repository")
    check_deidentification_and_ns(load_private_terms(args.private_terms_file))
    check_checklist()

    if args.out_dir:
        out_dir = pathlib.Path(args.out_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        run_e2e(out_dir)
    else:
        with tempfile.TemporaryDirectory() as tmp:
            run_e2e(pathlib.Path(tmp))

    check_repo_hygiene()
    print("\nRELEASE_GATE=ALL_PASS")
    return 0


def run_e2e(out_dir: pathlib.Path) -> None:
    # Audit shipped master directly
    out_m = _run_cli(
        [
            "audit",
            "--docx", str(MASTER_DOCX),
            "--mode", "phd-thesis",
            "--review-stage", "final",
            "--json-report", str(out_dir / "master_audit.json"),
        ],
        "E2E-1: Audit master docx",
    )
    assert "AUDIT_SUMMARY=FAIL:0" in out_m

    # Case A: blind PhD thesis
    blind_docx = out_dir / "test_phd_blind.docx"
    _run_cli(
        [
            "init",
            "--output", str(blind_docx),
            "--mode", "phd-thesis",
            "--review-stage", "blind",
            "--university", "示例大学",
            "--school", "生命科学学院",
            "--author", "测试作者",
            "--advisor", "测试导师",
            "--overwrite",
        ],
        "E2E-2a: Init phd-thesis blind",
    )
    out_a = _run_cli(
        [
            "audit",
            "--docx", str(blind_docx),
            "--mode", "phd-thesis",
            "--review-stage", "blind",
            "--sensitive-names", "测试作者,测试导师",
        ],
        "E2E-2b: Audit phd-thesis blind",
    )
    assert "AUDIT_SUMMARY=FAIL:0" in out_a

    # Case B: proposal + GB/T 7714 + custom JSON preset
    prop_docx = out_dir / "test_proposal_gb7714.docx"
    _run_cli(
        [
            "init",
            "--output", str(prop_docx),
            "--mode", "proposal",
            "--review-stage", "final",
            "--preset", "cn-gb7714-general",
            "--profile-config", str(CUSTOM_PRESET),
            "--citation-style", "gb7714",
            "--overwrite",
        ],
        "E2E-3a: Init proposal GB/T 7714 + custom preset",
    )
    out_b = _run_cli(
        [
            "audit",
            "--docx", str(prop_docx),
            "--mode", "proposal",
            "--preset", "cn-gb7714-general",
            "--profile-config", str(CUSTOM_PRESET),
            "--citation-style", "gb7714",
        ],
        "E2E-3b: Audit proposal GB/T 7714",
    )
    assert "AUDIT_SUMMARY=FAIL:0" in out_b

    # Case C: midterm & rotation quick init+audit
    for mode_name in ("midterm", "rotation"):
        m_docx = out_dir / f"test_{mode_name}.docx"
        _run_cli(["init", "--output", str(m_docx), "--mode", mode_name, "--overwrite"], f"E2E-4a: Init {mode_name}")
        out_c = _run_cli(["audit", "--docx", str(m_docx), "--mode", mode_name], f"E2E-4b: Audit {mode_name}")
        assert "AUDIT_SUMMARY=FAIL:0" in out_c

    # Case D: custom logo PNG embedding + export-checklist + list-presets
    logo_png = out_dir / "test_emblem.png"
    _make_1x1_png(logo_png)
    logo_docx = out_dir / "test_master_with_logo.docx"
    _run_cli(
        [
            "init",
            "--output", str(logo_docx),
            "--mode", "master-thesis",
            "--logo", str(logo_png),
            "--overwrite",
        ],
        "E2E-5a: Init master-thesis with custom PNG logo",
    )
    out_d = _run_cli(["audit", "--docx", str(logo_docx), "--mode", "master-thesis"], "E2E-5b: Audit master-thesis with custom logo")
    assert "AUDIT_SUMMARY=FAIL:0" in out_d
    with zipfile.ZipFile(logo_docx, "r") as zf:
        assert "word/media/logo.png" in zf.namelist()
        doc_xml = zf.read("word/document.xml").decode("utf-8")
        assert "ns0:" not in doc_xml
        assert "wp:inline" in doc_xml

    chk_copy = out_dir / "exported_checklist.docx"
    _run_cli(["export-checklist", "--output", str(chk_copy), "--overwrite"], "E2E-6: export-checklist")
    assert chk_copy.exists() and chk_copy.stat().st_size == CHECKLIST_DOCX.stat().st_size

    presets_out = _run_cli(["list-presets"], "E2E-7: list-presets")
    assert "cn-lifesci-default" in presets_out and "cn-gb7714-general" in presets_out


if __name__ == "__main__":
    raise SystemExit(main())
