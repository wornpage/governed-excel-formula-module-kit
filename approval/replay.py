"""Replay the five published examples from Excel's saved results. Python 3.10+."""
from hashlib import sha256
from pathlib import Path
from posixpath import join, normpath
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile
import json
import sys

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))
from formula_modules import alias_problems, changes, read_baseline, read_workbook

VERSIONS = ("A", "B-data", "B-formula", "B-expected", "B-combined")
# Text-file hash after Git newline normalization, from the public version A revision.
FIXTURE_SHA256 = "b92bb98455fb0e944a0f979c9a0490414e5a4b223154b36692fccf7b73e8796d"


def trusted_cases():
    fixture = (ROOT / "verification.json").read_text(encoding="utf-8")
    if sha256(fixture.encode()).hexdigest() != FIXTURE_SHA256:
        raise ValueError("The external version A fixture changed.")
    rows = json.loads(fixture)["reopenedCases"]
    if [r["row"] for r in rows] != list(range(6, 14)):
        raise ValueError("Expected the fixed eight cases, rows 6 through 13.")
    return rows


def cells(path):
    """Read entered values, saved formulas and cached results for the two example sheets."""
    with ZipFile(path) as z:
        strings = ["".join(t.text or "" for t in si.findall(".//{*}t"))
                   for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = {r.attrib["Id"]: r.attrib["Target"]
                for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        sheets = {}
        for sheet in wb.findall("{*}sheets/{*}sheet"):
            rid = next(v for k, v in sheet.attrib.items() if k.endswith("}id"))
            target = rels[rid]
            part = target[1:] if target.startswith("/") else normpath(join("xl", target))
            data = {}
            for c in ET.fromstring(z.read(part)).findall(".//{*}sheetData/{*}row/{*}c"):
                f, v = c.find("{*}f"), c.find("{*}v")
                if f is None and v is None:  # format-only cell
                    continue
                typ = c.get("t", "n")
                value = "" if v is None or v.text is None else v.text
                if typ == "s":
                    typ, value = "str", strings[int(value)]
                formula = None if f is None else (f.text or "", tuple(sorted(f.attrib.items())))
                data[c.attrib["r"]] = (typ, value, formula)
            sheets[sheet.attrib["name"]] = data
    if set(sheets) != {"Cases", "Records"}:
        raise ValueError("Unexpected worksheets.")
    return sheets


def authored(cell):
    # A recalculated result is an effect of an edit, not another authored edit.
    return None if cell is None else ("formula", cell[2]) if cell[2] is not None else cell[:2]


def evidence(path, fixture, baseline, a_cells):
    names, source = read_workbook(path)
    data = cells(path)
    edits = sorted(f"{sheet}!{ref}" for sheet, vals in data.items()
                   for ref in vals.keys() | a_cells[sheet].keys()
                   if authored(vals.get(ref)) != authored(a_cells[sheet].get(ref)))
    sheet_fail, result_fail, expected_edit = [], [], []
    for case in fixture:
        n, expected = case["row"], case["expected"]
        row = data["Cases"]
        if row[f"B{n}"][1] != case["case"]:
            raise ValueError(f"Unexpected identity for case {n}.")
        actual_type, actual, _ = row[f"E{n}"]
        entered = row.get(f"D{n}")
        mark = row[f"F{n}"][1]
        if mark not in ("PASS", "FAIL"):
            raise ValueError(f"No saved PASS/FAIL at Cases!F{n}.")
        if mark == "FAIL":
            sheet_fail.append(n)
        if actual_type != "str" or row[f"E{n}"][2] is None or actual != expected:
            result_fail.append({"row": n, "approved": expected, "returned": actual})
        if (entered is None and expected != "") or (entered is not None and
                (entered[0] != "str" or entered[1] != expected or entered[2] is not None)):
            expected_edit.append({"row": n, "approved": expected,
                                  "entered": "" if entered is None else entered[1]})
    return {"sha256": source["sha256"],
            "definitionChanges": changes(baseline, names) + alias_problems(names),
            "editedWorksheetCells": edits, "caseCount": len(fixture),
            "sheetFailedRows": sheet_fail, "returnedValueMismatches": result_fail,
            "expectedCellMismatches": expected_edit}


def main():
    receipt = json.loads((HERE / "decisions.json").read_text(encoding="utf-8"))
    baseline, source = read_baseline(ROOT / "modules")
    fixture = trusted_cases()
    a_cells = cells(ROOT / "example.xlsx")
    print("Version     Names     Sheet  Returned  Expected cells   Illustrative decision")
    for version in VERSIONS:
        path = ROOT / "example.xlsx" if version == "A" else HERE / (version + ".xlsx")
        decision = receipt["versions"][version]
        found = evidence(path, fixture, baseline, a_cells)
        if found["sha256"] != decision["evidence"]["sha256"]:
            raise ValueError(f"{version}: workbook bytes changed; its recorded decision does not apply.")
        if (version == "A" and found["sha256"] != source["sha256"]) or found != decision["evidence"]:
            raise ValueError(f"{version}: saved evidence differs from the example receipt.")
        n = found["caseCount"]
        scores = [f"{n-len(found[key])}/{n}" for key in
                  ("sheetFailedRows", "returnedValueMismatches", "expectedCellMismatches")]
        names = "changed" if found["definitionChanges"] else "match"
        print(f"{version:<11} {names:<9} {scores[0]:<6} {scores[1]:<9} {scores[2]:<16} {decision['decision']}")
    print("Matched the published examples. This reads saved values; it does not recalculate or grant approval.")


if __name__ == "__main__":
    try:
        main()
    except (BadZipFile, OSError, ET.ParseError, KeyError, ValueError, IndexError, StopIteration) as exc:
        sys.exit("Cannot replay: " + str(exc))
