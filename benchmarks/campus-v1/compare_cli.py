"""Run two pinned, unmodified command-line tools on the same synthetic Word inputs."""
import argparse
import csv
import json
from pathlib import Path
import subprocess
import sys


def invoke(script, args, allowed=(0,)):
    run = subprocess.run(
        [sys.executable, "-B", "-X", "utf8", str(script), *map(str, args)],
        capture_output=True, encoding="utf-8", timeout=90, check=False,
    )
    if run.returncode not in allowed:
        raise RuntimeError(f"{script.name} exited {run.returncode}: {run.stderr[-1200:]}")
    return run.returncode


def ours_labels(report):
    by_id = {row["check_id"]: row for row in report["checks"]}
    labels = []
    for rule, label in [
        ("Layout-1", "PAGE"), ("Layout-SecNum", "NUMBER"), ("Layout-TOC", "TOC"),
        ("Pkg-Media", "MEDIA"), ("Acad-10-Blind", "IDENTITY"),
    ]:
        if by_id[rule]["status"] == "FAIL":
            labels.append(label)
    styles = by_id["Layout-TemplateStyles"]
    if styles["status"] == "WARN" and "BodyCampus" in styles["details"]:
        labels.append("STYLE")
    return sorted(labels)


def upstream_labels(rows):
    page = {f"section.margin.{side}" for side in ("left", "right", "top", "bottom")}
    page.update({"section.page.width", "section.page.height", "section.page.orientation"})
    labels = set()
    for row in rows:
        if row["status"] != "fail":
            continue
        if row["category"] == "section" and row["property"] in page:
            labels.add("PAGE")
        if row["target"] == "style:BodyCampus" and row["property"] == "font.size_half_points":
            labels.add("STYLE")
    return sorted(labels)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    for arg in ("inputs", "plugin", "upstream", "out"):
        p.add_argument("--" + arg, required=True, type=Path)
    a = p.parse_args()
    a.out.mkdir(parents=True, exist_ok=False)
    cli = a.plugin / "skills/bio-thesis-formatter/scripts/bio_thesis.py"
    scripts = a.upstream / "scripts"
    school = a.inputs / "school.docx"
    requirements = a.out / "requirements"
    invoke(scripts / "extract_requirements.py", ["--template", school, "--out", requirements])
    tools = ["v020", "format-thesis-docx"]
    results = {tool: {} for tool in tools}
    for tool in tools:
        (a.out / tool).mkdir()
    for i in range(1, 9):
        case = f"A{i:02d}"
        candidate = a.inputs / (case + ".docx")
        report = a.out / tools[0] / (case + ".json")
        invoke(cli, ["audit", "--docx", candidate, "--template", school, "--base-docx", school,
                     "--review-stage", "blind", "--sensitive-names", "EXAMPLE-Q17",
                     "--json-report", report], allowed=(0, 1))
        results[tools[0]][case] = ours_labels(json.loads(report.read_text(encoding="utf-8")))
        out = a.out / tools[1] / case
        invoke(scripts / "audit_format.py", ["--requirements", requirements / "requirements.json",
                                             "--paper", candidate, "--out", out])
        with (out / "format_audit.csv").open(encoding="utf-8-sig", newline="") as handle:
            results[tools[1]][case] = upstream_labels(csv.DictReader(handle))
    invoke(cli, ["init", "--template", a.inputs / "school.dotx",
                 "--author", "EXAMPLE-Q17", "--review-stage", "blind",
                 "--title-en", "Forms of Evidence", "--output", a.out / tools[0] / "generated.docx"])
    for tool, cases in results.items():
        (a.out / tool / "result.json").write_text(json.dumps({"cases": cases}, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
