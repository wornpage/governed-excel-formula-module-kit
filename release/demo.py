"""Replay the release owner's three decisions in a temporary local folder. Python 3.10+."""
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile
import json
import sys

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "approval"))
from replay import cells, evidence, read_baseline, trusted_cases

REVIEWER = "Riley"
RELEASE_OWNER = "Morgan"


@dataclass(frozen=True)
class Decision:
    status: str
    reviewed_sha256: str
    reviewer: str = REVIEWER


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest(snapshot):
    return sha256(snapshot).hexdigest()


def publish(actor, decision, submission, current):
    """Only the example release owner can replace Current.xlsx with the reviewed snapshot."""
    if actor != RELEASE_OWNER:
        raise PermissionError("Only the designated release owner can publish.")
    if decision.reviewer != REVIEWER or decision.status not in ("ACCEPT", "HOLD", "REJECT"):
        raise ValueError("Unrecognized review decision.")
    if decision.status != "ACCEPT":
        return decision.status
    snapshot = submission.read_bytes()
    if digest(snapshot) != decision.reviewed_sha256:
        return "STALE"
    # Write the bytes we checked. Reopening the submission for copying could read C instead.
    tmp = NamedTemporaryFile(dir=current.parent, delete=False)
    staged = Path(tmp.name)
    try:
        with tmp:
            tmp.write(snapshot)
        staged.replace(current)
    finally:
        staged.unlink(missing_ok=True)
    return "RELEASED"


def checked_inputs():
    versions = json.loads((ROOT / "approval" / "decisions.json").read_text(encoding="utf-8"))["versions"]
    extra = json.loads((HERE / "evidence.json").read_text(encoding="utf-8"))["workbooks"]
    a = ROOT / "example.xlsx"
    b = ROOT / "approval" / "B-data.xlsx"
    c = HERE / "C-replacement.xlsx"
    mixed = HERE / "B-mixed.xlsx"
    fixture = trusted_cases()
    baseline, source = read_baseline(ROOT / "modules")
    original = cells(a)
    require(digest(a.read_bytes()) == source["sha256"], "Version A differs from its snapshot.")
    for path, record, sheets in (
        (a, versions["A"], ("Cases", "Records")),
        (b, versions["B-data"], ("Cases", "Records")),
        (c, extra["C-replacement"], ("Cases", "Records")),
        (mixed, extra["B-mixed"], ("Cases", "Records", "Summary")),
    ):
        result = evidence(path, fixture, baseline, original, sheets)
        require(result == record["evidence"], f"Changed workbook or receipt: {path.name}")
        require(result["caseCount"] == 8 and not any(result[k] for k in (
            "definitionChanges", "sheetFailedRows", "returnedValueMismatches", "expectedCellMismatches")),
            f"The eight original checks or the named definition changed: {path.name}")
    # Use the previous review's two focused rename probes in addition to the original eight.
    probes = versions["B-data"]["additionalExcelProbes"]
    expected_probes = [("Label A", "", ""), ("Label B", "ID-108", "ID-108")]
    require([(p["selectedLabel"], p["expected"], p["returned"]) for p in probes] == expected_probes
            and versions["B-data"]["decision"] == "Accept this B", "B lacks its reviewed rename evidence.")
    mixed_cells = cells(mixed, ("Cases", "Records", "Summary"))
    edits = extra["B-mixed"]["evidence"]["editedWorksheetCells"]
    changed_callers = []
    for address in edits:
        sheet, ref = address.split("!")
        cell = mixed_cells[sheet].get(ref)
        if cell is not None and cell[2] is not None:
            changed_callers.append(address)
    require("Records!B13" in edits and changed_callers == ["Summary!C6"], "Expected an untested mixed edit.")
    return a, b, c, mixed, changed_callers, mixed_cells["Summary"], fixture


def perform(current, submission, decision, expected_status, expected_current):
    status = publish(RELEASE_OWNER, decision, submission, current)
    require(status == expected_status and current.read_bytes() == expected_current.read_bytes(),
            f"Unexpected release result: {status}, {current.name}")
    return status


def main():
    a, b, c, mixed, uncovered, summary, fixture = checked_inputs()
    with TemporaryDirectory() as folder:
        inbox = Path(folder) / "inbox"
        released = Path(folder) / "released"
        inbox.mkdir()
        released.mkdir()
        submission = inbox / "Submission.xlsx"
        current = released / "Current.xlsx"

        # Each scenario starts with A already current.
        current.write_bytes(a.read_bytes())
        submission.write_bytes(b.read_bytes())
        b_request = digest(submission.read_bytes())
        perform(current, submission, Decision("ACCEPT", b_request), "RELEASED", b)
        print("1. RELEASED | Morgan published the B that Riley reviewed | current B")

        current.write_bytes(a.read_bytes())
        submission.write_bytes(b.read_bytes())
        pending_b = digest(submission.read_bytes())       # Riley opens B for review.
        submission.write_bytes(c.read_bytes())            # The same submission link now serves C.
        perform(current, submission, Decision("ACCEPT", pending_b), "STALE", a)
        print("2. STALE    | Riley's response covers B, but Morgan found C | current A")

        current.write_bytes(a.read_bytes())
        submission.write_bytes(mixed.read_bytes())
        mixed_request = digest(submission.read_bytes())
        if uncovered:
            perform(current, submission, Decision("HOLD", mixed_request), "HOLD", a)
            print(f"3. HOLD     | 8/8 original cases pass; {uncovered[0]} needs its own check | current A")
        beta = next(row for row in fixture if row["row"] == 7 and row["case"] == "Two matching labels")
        cell_type, returned, _ = summary["C6"]
        require(summary["B6"][1] == "Beta" and cell_type == "str"
                and beta["expected"] == "" and returned == "ID-102", "Unexpected focused caller result.")
        if returned != beta["expected"]:
            perform(current, submission, Decision("REJECT", mixed_request), "REJECT", a)
            print("   REJECT   | additional caller check 0/1: Beta returned ID-102 instead of empty text | current A")


if __name__ == "__main__":
    try:
        main()
    except (OSError, BadZipFile, ParseError, KeyError, ValueError, IndexError, StopIteration) as exc:
        sys.exit(f"Cannot replay the example: {exc}")
