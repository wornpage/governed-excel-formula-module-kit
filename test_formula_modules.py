"""Regression tests for the workbook reader and snapshot CLI. No Excel installation required."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile


SCRIPT = Path(__file__).with_name("formula_modules.py")
LAMBDA = '_xlfn.LAMBDA(_xlpm.v,IF(_xlpm.v="a b","a  b",0))'
ALIAS = {"comment":"Governed formula compatibility alias: kind"}
ROWS = [
    ("kind.Core", LAMBDA, {}), ("Core", LAMBDA, ALIAS),
    ("Controls.Mode", '"All"', {}),
    ("Mode", "'Data'!$C$2", {"comment":"Governed formula visible control: Data!C2"}),
    ("get.First.Second", "1", {"hidden":"1"}),
    ("_xlnm.Print_Area", "Sheet!$A$1", {"localSheetId":"0"}),
]


def save_package(path, rows, marker=""):
    root = ET.Element("workbook", xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    sheets = ET.SubElement(root, "sheets")
    for sheet in ("Workbook", "Data"):
        ET.SubElement(sheets, "sheet", name=sheet)
    names = ET.SubElement(root, "definedNames")
    for name, formula, attrs in rows:
        ET.SubElement(names, "definedName", name=name, **attrs).text = formula
    with ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", ET.tostring(root, encoding="utf-8"))
        z.writestr("docProps/test.txt", marker)


def run_cli(*args):
    return subprocess.run([sys.executable, "-B", str(SCRIPT), *map(str, args)], capture_output=True, text=True)


def bytes_by_name(path):
    return {f.name:f.read_bytes() for f in path.iterdir()}


class ModuleTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.book = self.root / "example.xlsx"
        self.baseline = self.root / "approved"
        save_package(self.book, ROWS)
        exported = run_cli("export", self.book, self.baseline)
        self.assertEqual(exported.returncode, 0, exported.stderr)

    def check(self):
        return run_cli("check", self.book, "--baseline", self.baseline)

    def test_repeatable_export_and_read_only_check(self):
        initial = bytes_by_name(self.baseline)
        replay = self.root / "replay"
        self.assertEqual(run_cli("export", self.book, replay).returncode, 0)
        self.assertEqual(initial, bytes_by_name(replay))
        self.assertEqual(run_cli("export", self.book, self.baseline).returncode, 0)
        manifest = json.loads(initial["manifest.json"])
        self.assertEqual(manifest["source"]["file"], self.book.name)
        self.assertEqual(manifest["counts"]["defined_names"], 5)
        self.assertEqual(manifest["counts"]["declared_aliases"], 1)
        self.assertEqual(manifest["counts"]["visible_controls"], 1)
        get, = json.loads(initial["module-get.json"])
        self.assertEqual((get["name"], get["hidden"]), ("get.First.Second", True))
        kind, = json.loads(initial["module-kind.json"])
        self.assertEqual(kind["formula"], LAMBDA)
        # A different package fingerprint and XML order do not change the governed formulas.
        save_package(self.book, reversed(ROWS), marker="resaved")
        book_bytes = self.book.read_bytes()
        checked = self.check()
        self.assertEqual(checked.returncode, 0, checked.stderr)
        self.assertIn("bytes differ", checked.stdout)
        self.assertEqual(self.book.read_bytes(), book_bytes)
        self.assertEqual(initial, bytes_by_name(self.baseline))
        backlog = run_cli("check", self.book, "--baseline", self.baseline, "--list-uncommented")
        self.assertEqual(backlog.returncode, 0)
        self.assertIn("NO COMMENT [workbook]!get.First.Second", backlog.stdout)
        self.assertNotIn("NO COMMENT [workbook]!Mode", backlog.stdout)

    def test_native_comment_line_breaks_and_literal_escape(self):
        rows = ROWS.copy()
        raw = "Governed formula compatibility alias: kind_x000a_Keep _x005F_x000a_ literal. _xD83D__xDE80_"
        shown = "Governed formula compatibility alias: kind\nKeep _x000a_ literal. 🚀"
        rows[1] = ("Core", LAMBDA, {"comment": raw})
        save_package(self.book, rows)
        candidate = self.root / "multiline"
        exported = run_cli("export", self.book, candidate)
        self.assertEqual(exported.returncode, 0, exported.stderr)
        names = json.loads((candidate / "unprefixed.json").read_text(encoding="utf-8"))
        alias = next(item for item in names if item["name"] == "Core")
        self.assertEqual(alias["comment"], shown)
        self.assertEqual(alias["formula"], LAMBDA)
        self.assertEqual(run_cli("check", self.book, "--baseline", candidate).returncode, 0)

    def test_check_prior_snapshot_with_excel_comment_escapes(self):
        rows = ROWS.copy()
        rows[0] = ("kind.Core", LAMBDA, {"comment":"Line one_x000a_Line two"})
        rows[4] = ("get.First.Second", "1", {"hidden":"1", "comment":"_x0020_"})
        save_package(self.book, rows)

        manifest_path = self.baseline / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["schema"] = 1
        manifest["source"]["sha256"] = hashlib.sha256(self.book.read_bytes()).hexdigest()
        manifest["counts"]["without_comment"] = 1
        for filename, name, comment in (
            ("module-kind.json", "kind.Core", "Line one_x000a_Line two"),
            ("module-get.json", "get.First.Second", "_x0020_"),
        ):
            path = self.baseline / filename
            records = json.loads(path.read_text(encoding="utf-8"))
            next(record for record in records if record["name"] == name)["comment"] = comment
            contents = (json.dumps(records, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
            path.write_bytes(contents)
            manifest["modules"][filename] = hashlib.sha256(contents).hexdigest()
        manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")

        checked = self.check()
        self.assertEqual(checked.returncode, 0, checked.stdout + checked.stderr)
        repeated = run_cli("export", self.book, self.baseline)
        self.assertEqual(repeated.returncode, 0, repeated.stdout + repeated.stderr)
        rows[0] = ("kind.Core", LAMBDA, {"comment":"A changed comment"})
        save_package(self.book, rows)
        changed = self.check()
        self.assertEqual(changed.returncode, 1)
        self.assertIn("CHANGED [workbook]!kind.Core: comment", changed.stdout)
        self.assertNotIn("CHANGED [workbook]!get.First.Second", changed.stdout)

    def test_added_removed_changed_and_alias_drift(self):
        rows = ROWS.copy()
        rows[1] = ("Core", LAMBDA.replace('"a  b"', '"a b"'), ALIAS)
        rows[3] = ("Mode", rows[3][1], {"comment":"new note","hidden":"true"})
        rows[4] = ("get.New", "1", {})
        save_package(self.book, rows)
        checked = self.check()
        self.assertEqual(checked.returncode, 1)
        for line in ("ADDED [workbook]!get.New", "REMOVED [workbook]!get.First.Second",
                     "CHANGED [workbook]!Core: formula", "CHANGED [workbook]!Mode: comment, hidden",
                     "ALIAS DRIFT [workbook]!Core -> [workbook]!kind.Core"):
            self.assertIn(line, checked.stdout)
        candidate = self.root / "candidate"
        self.assertEqual(run_cli("export", self.book, candidate).returncode, 1)
        self.assertFalse(candidate.exists())

    def test_sheet_scope_case_matching_and_missing_target(self):
        rows = ROWS + [("core", "2", dict(ALIAS, localSheetId="0")),
                       ("KIND.Core", "2", {"localSheetId":"0"}),
                       ("Core", LAMBDA, dict(ALIAS, localSheetId="1"))]
        save_package(self.book, rows)
        scoped = self.root / "scoped"
        self.assertEqual(run_cli("export", self.book, scoped).returncode, 0)
        unprefixed = json.loads((scoped / "unprefixed.json").read_text())
        self.assertEqual({n["scope"] for n in unprefixed if n["name"].lower()=="core"}, {None,"Workbook","Data"})
        removed = [r for r in rows if r[0] != "kind.Core"]
        save_package(self.book, removed)
        checked = run_cli("check", self.book, "--baseline", scoped)
        self.assertEqual(checked.returncode, 1)
        self.assertIn("ALIAS MISSING TARGET Data!Core", checked.stdout)
        self.assertIn("ALIAS MISSING TARGET [workbook]!Core", checked.stdout)
        self.assertNotIn("ALIAS MISSING TARGET Workbook!core", checked.stdout)

    def test_invalid_scopes_duplicates_and_bad_manifest(self):
        for extra in [("CORE", "1", {}), ("test.Bad", "1", {"localSheetId":"-1"})]:
            save_package(self.book, ROWS + [extra])
            self.assertEqual(self.check().returncode, 2)
        save_package(self.book, ROWS)
        path = self.baseline / "module-get.json"
        path.write_text("[]\n", encoding="utf-8")
        checked = self.check()
        self.assertEqual(checked.returncode, 2)
        self.assertIn("checksum differs", checked.stderr)

    def test_export_preserves_an_existing_different_snapshot(self):
        initial = bytes_by_name(self.baseline)
        save_package(self.book, ROWS + [("get.New", "123", {})])
        result = run_cli("export", self.book, self.baseline)
        self.assertEqual(result.returncode, 2)
        self.assertIn("different snapshot", result.stderr)
        self.assertEqual(initial, bytes_by_name(self.baseline))

    def test_intentional_change_to_alias_and_target_can_be_reviewed(self):
        rows = ROWS.copy()
        revised = LAMBDA.replace('"a  b"', '"approved"')
        rows[:2] = [("kind.Core", revised, {}), ("Core", revised, ALIAS)]
        save_package(self.book, rows)
        before = self.check()
        self.assertEqual(before.returncode, 1)
        self.assertIn("CHANGED [workbook]!kind.Core: formula", before.stdout)
        self.assertIn("CHANGED [workbook]!Core: formula", before.stdout)
        self.assertNotIn("ALIAS DRIFT", before.stdout)
        candidate = self.root / "candidate"
        self.assertEqual(run_cli("export", self.book, candidate).returncode, 0)
        self.assertEqual(run_cli("check", self.book, "--baseline", candidate).returncode, 0)


if __name__ == "__main__":
    unittest.main()
