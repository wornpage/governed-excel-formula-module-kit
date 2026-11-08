"""Export and compare the saved defined names of an Excel workbook. Python 3.10+."""

import argparse
from dataclasses import asdict, dataclass, replace
import hashlib
from io import BytesIO
import json
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from urllib.parse import quote
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


DEFAULT_BASELINE = Path(__file__).resolve().parent / "modules"
ALIAS_TAG = "Governed formula compatibility alias:"
CONTROL_TAG = "Governed formula visible control:"
DIRECT_LAMBDA = re.compile(r"^\s*=?\s*(?:_xlfn\.)?LAMBDA\s*\(", re.I)
WORKBOOK_NAMESPACES = {
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
}


@dataclass(frozen=True)
class Name:
    name: str
    scope: str | None  # None = workbook. A sheet can itself be named "Workbook".
    formula: str
    comment: str
    hidden: bool

    def __post_init__(self):
        if not isinstance(self.name, str) or not self.name:
            raise ValueError("A defined name must have a nonempty name")
        if self.scope is not None and (not isinstance(self.scope, str) or not self.scope):
            raise ValueError(f"Invalid sheet scope for {self.name}")
        if not isinstance(self.formula, str) or not isinstance(self.comment, str) or type(self.hidden) is not bool:
            raise ValueError(f"Invalid formula, comment or hidden flag for {self.name}")

    @property
    def key(self):
        sheet = self.scope.casefold() if self.scope is not None else None
        return sheet, self.name.casefold()

    @property
    def label(self):
        return f"{self.scope or '[workbook]'}!{self.name}"


def json_text(value):
    return json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False) + "\n"


def by_name(names):
    index = {}
    for item in names:
        if item.key in index:
            raise ValueError(f"Duplicate defined name in the same scope: {item.label}")
        index[item.key] = item
    return index


def read_comment(value):
    # SpreadsheetML uses one-pass UTF-16 escapes, including _x005F_ for a literal underscore.
    if "_x" not in value:
        return value
    decoded = re.sub(r"_x([0-9a-fA-F]{4})_", lambda match: chr(int(match[1], 16)), value)
    return decoded.encode("utf-16-le", errors="surrogatepass").decode("utf-16-le")


def read_workbook(path):
    # Hash and inspect the very same bytes, including when the live file is replaced.
    data = path.read_bytes()
    with ZipFile(BytesIO(data)) as package:
        root = ET.fromstring(package.read("xl/workbook.xml"))
    if root.tag not in {f"{{{ns}}}workbook" for ns in WORKBOOK_NAMESPACES}:
        raise ValueError("Expected a SpreadsheetML workbook")
    sheets = [sheet.attrib["name"] for sheet in root.findall("{*}sheets/{*}sheet")]
    names = []
    for node in root.findall("{*}definedNames/{*}definedName"):
        full = node.attrib["name"]
        if full.casefold().startswith("_xlnm."):
            continue
        scope = None
        if "localSheetId" in node.attrib:
            sheet_id = int(node.attrib["localSheetId"])
            if not 0 <= sheet_id < len(sheets):
                raise ValueError(f"Invalid sheet scope for {full}: {sheet_id}")
            scope = sheets[sheet_id]
        hidden = node.get("hidden", "0").lower()
        if hidden not in {"0", "1", "false", "true"}:
            raise ValueError(f"Invalid defined-name metadata: {full!r}")
        names.append(Name(full, scope, node.text or "", read_comment(node.get("comment", "")), hidden in {"1", "true"}))
    return by_name(names), {"file": path.name, "sha256": hashlib.sha256(data).hexdigest()}


def module_file(name):
    prefix, dot, _ = name.partition(".")
    if not dot:
        return "unprefixed.json"
    # Use one file for all case variants. Encoding also keeps workbook text out of paths.
    slug = quote(prefix.casefold(), safe="")
    if len(slug) > 140:
        return "long-module-" + hashlib.sha256(slug.encode("ascii")).hexdigest() + ".json"
    return f"module-{slug}.json"


def tag_line(item):
    return item.comment.partition("\n")[0].strip()


def alias_problems(names):
    problems = []
    for item in names.values():
        line = tag_line(item)
        if not line.casefold().startswith(ALIAS_TAG.casefold()):
            continue
        module = line[len(ALIAS_TAG):].strip()
        if not module or "." in module or "." in item.name:
            problems.append(f"ALIAS TAG {item.label}: expected an unprefixed name and one target module")
            continue
        target_name = f"{module}.{item.name}".casefold()
        target = names.get((item.key[0], target_name)) or names.get((None, target_name))
        if target is None:
            problems.append(f"ALIAS MISSING TARGET {item.label} -> {module}.{item.name}")
        elif item.formula != target.formula:
            problems.append(f"ALIAS DRIFT {item.label} -> {target.label}: formula differs")
    return sorted(problems, key=str.casefold)


def counts(names):
    rows = names.values()
    return {
        "defined_names": len(names),
        "direct_lambdas": sum(bool(DIRECT_LAMBDA.match(n.formula)) for n in rows),
        "declared_aliases": sum(tag_line(n).casefold().startswith(ALIAS_TAG.casefold()) for n in rows),
        "visible_controls": sum(tag_line(n).casefold().startswith(CONTROL_TAG.casefold()) for n in rows),
        "without_comment": sum(not n.comment.strip() for n in rows),
    }


def snapshot_files(names, source):
    groups = {}
    for item in sorted(names.values(), key=lambda n: (n.name.casefold(), (n.scope or "").casefold())):
        groups.setdefault(module_file(item.name), []).append(asdict(item))
    files = {filename: json_text(groups[filename]) for filename in sorted(groups)}
    hashes = {filename: hashlib.sha256(content.encode("utf-8")).hexdigest() for filename, content in files.items()}
    files["manifest.json"] = json_text({"schema": 2, "source": source, "counts": counts(names), "modules": hashes})
    return files


def read_baseline(directory):
    manifest_path = directory / "manifest.json"
    if manifest_path.is_symlink() or not manifest_path.is_file():
        raise ValueError(f"Missing regular manifest in {directory}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or manifest.get("schema") not in (1, 2):
        raise ValueError("Unsupported formula-module snapshot format")
    modules = manifest["modules"]
    source = manifest["source"]
    if (not isinstance(modules, dict) or not isinstance(source, dict) or not isinstance(source.get("file"), str)
            or not re.fullmatch(r"[a-f0-9]{64}", source.get("sha256", ""))):
        raise ValueError("Invalid source or module index in manifest")
    # Enumerate actual children; never open paths supplied inside a manifest.
    files = {f.name: f for f in directory.iterdir() if f.name != "manifest.json"}
    if files.keys() != modules.keys():
        raise ValueError("Baseline file list differs from its manifest")
    names = []
    for filename, path in files.items():
        if path.is_symlink() or not path.is_file():
            raise ValueError(f"Expected a regular module file: {filename}")
        text = path.read_text(encoding="utf-8")  # Normalize CRLF/LF used by Git checkouts.
        if hashlib.sha256(text.encode("utf-8")).hexdigest() != modules[filename]:
            raise ValueError(f"Baseline checksum differs for {filename}; regenerate the snapshot")
        records = json.loads(text)
        if not isinstance(records, list) or not records:
            raise ValueError(f"Expected named definitions in {filename}")
        for record in records:
            item = Name(**record)
            if module_file(item.name) != filename or item.name.casefold().startswith("_xlnm."):
                raise ValueError(f"Unexpected name in {filename}: {item.name}")
            names.append(item)
    result = by_name(names)
    if manifest["counts"] != counts(result):
        raise ValueError("Baseline counts differ from the module files")
    if manifest["schema"] == 1:
        result = by_name(replace(item, comment=read_comment(item.comment)) for item in names)
    return result, source


def export_snapshot(names, source, output):
    files = snapshot_files(names, source)
    output.parent.mkdir(parents=True, exist_ok=True)
    # An export is a candidate for review. Publish a complete directory in one rename.
    with TemporaryDirectory(dir=output.parent, prefix=".formula-export-") as scratch:
        candidate = Path(scratch) / "modules"
        candidate.mkdir()
        for filename, text in files.items():
            (candidate / filename).write_bytes(text.encode("utf-8"))
        try:
            candidate.rename(output)
        except OSError:
            if output.is_symlink() or not output.is_dir():
                raise
            current, original_source = read_baseline(output)
            if current != names or original_source != source:
                raise ValueError("Output contains a different snapshot. Export the candidate to a new directory.")


def changes(baseline, current):
    messages = []
    for key, old in baseline.items():
        if key not in current:
            messages.append(f"REMOVED {old.label}")
        else:
            new = current[key]
            fields = [field for field in ("name", "scope", "formula", "comment", "hidden")
                      if getattr(old, field) != getattr(new, field)]
            if fields:
                messages.append(f"CHANGED {new.label}: {', '.join(fields)}")
    messages.extend(f"ADDED {item.label}" for key, item in current.items() if key not in baseline)
    return sorted(messages, key=str.casefold)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    check = commands.add_parser("check", help="compare a workbook to the approved modules, without writing files")
    check.add_argument("workbook", type=Path)
    check.add_argument("--baseline", type=Path, default=DEFAULT_BASELINE)
    check.add_argument("--list-uncommented", action="store_true", help="list the current documentation queue")
    export = commands.add_parser("export", help="write a deterministic candidate snapshot")
    export.add_argument("workbook", type=Path)
    export.add_argument("output", type=Path, help="new snapshot directory")
    args = parser.parse_args()
    names, source = read_workbook(args.workbook)
    issues = alias_problems(names)
    if args.command == "check":
        baseline, original = read_baseline(args.baseline)
        issues = changes(baseline, names) + issues
        if source["sha256"] != original["sha256"]:
            print("Workbook bytes differ from the recorded source. Comparing defined names.")
        if args.list_uncommented:
            for label in sorted((n.label for n in names.values() if not n.comment.strip()), key=str.casefold):
                print(f"NO COMMENT {label}")
    if issues:
        print("\n".join(issues))
        return 1
    if args.command == "export":
        export_snapshot(names, source, args.output)
        print(f"Exported {args.output}")
    summary = counts(names)
    print(f"PASS: {len(names)} names, {summary['direct_lambdas']} direct LAMBDAs, "
          f"{summary['declared_aliases']} matching aliases, {summary['visible_controls']} visible controls, "
          f"{summary['without_comment']} without comments.")
    print(f"Workbook SHA-256: {source['sha256']}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, ValueError, TypeError, KeyError, BadZipFile, ET.ParseError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(2)
