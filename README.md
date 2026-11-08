# Solving governed Excel formula modules

Start with a rule someone else can test. This kit shows how to version a named Excel formula and check its behavior using a table with two fields, a label and a record ID. Try it with the supplied data, then apply the same method to a workbook of your own.

[Download the complete kit (ZIP)](https://github.com/wornpage/governed-excel-formula-module-kit/archive/refs/heads/main.zip)

## Try it in Excel

1. Open [example.xlsx](example.xlsx) in Excel with LAMBDA support. The **Cases** sheet should show eight **PASS** results. Two rows on **Records** have the label **Beta**, one with ID-102 and the other with ID-103. The expected result is empty because the label alone cannot identify a single row.
2. On a copy, change **Records!B8** from **Beta** to **Theta**. Recalculate. **Cases!E7** should return **ID-102**, and **F7** should say **FAIL** because the expected result describes the original two matches. Undo the edit and the check should return to **PASS**.
3. On Windows, open **Formulas → Name Manager** and inspect `Lookup.ResolveID`. The name is scoped to the workbook. Its comment explains the inputs and return value. [Microsoft’s LAMBDA instructions](https://support.microsoft.com/en-us/excel/functions/lambda-function)

The function takes a selected label and two aligned, one-column ranges of labels and text IDs. It returns one ID when the label occurs once and its nonblank ID also occurs once. Otherwise it returns empty text. Matching ignores capitalization and treats an asterisk as a literal character. The function checks the first and last exact match. Different positions mean the value occurs more than once. [Microsoft’s XMATCH reference](https://support.microsoft.com/en-us/excel/functions/xmatch-function)

## Apply the method to your workbook

| Step | What the reviewer needs |
| --- | --- |
| 1. Define the rule | Record the full name and scope, inputs and units, expected value or spill dimensions, and what blanks, zeros, errors and duplicate keys mean. For a KPI, state its numerator, denominator, filters and period. |
| 2. Capture the baseline | Save a copy of the workbook. Export its named definitions into a new directory. The manifest identifies the workbook by its SHA-256 fingerprint. Keep this directory as the comparison point. |
| 3. Reproduce the failure | Use the smallest inputs that trigger the defect. Write the expected answer from the rule before changing the formula. Run the original formula and record its actual answer. Keep at least one ordinary case that already works. |
| 4. Fix the owning name | Follow the failing worksheet cell to the named formula and its dependencies. Change the name that owns the rule. Review affected names, aliases, worksheet formulas, validation rules and other consumers. |
| 5. Verify and hand off | Compare the candidate against the baseline. In the target Excel version, recalculate, save and reopen a working copy. Re-run the failing case and affected callers. Record expected and actual results, Excel build, workbook fingerprint and the reviewer’s decision. Preserve the prior workbook and snapshot for rollback. |

When a rule changes intentionally, record the old and new expectations. For example, the case pass rate uses all eight cases as its denominator. Removing a failing case changes that denominator. Keep the case visible and document any change to the set.

## Export and check

The checker uses Python 3.10 or later and its standard library. Run these commands from the extracted kit folder. If your Python executable is called `python3`, substitute it for `python`.

```text
python formula_modules.py check example.xlsx
```

A successful run reports one named LAMBDA. To use your own saved workbooks, substitute your filenames and choose new directories for each export.

```text
python formula_modules.py export "original.xlsx" approved
python formula_modules.py check "changed.xlsx" --baseline approved
python formula_modules.py export "changed.xlsx" candidate
```

Compare `approved` and `candidate` as text. Review the whole directory so removed names and modules stay visible. Keep the reviewed candidate snapshot with its matching workbook.

| Checker exit | Meaning |
| --- | --- |
| 0 | Governed definitions match on a check, or an export succeeded. A different workbook fingerprint is reported separately. |
| 1 | Names were added, removed or changed, or a declared alias disagrees with its target. An intended fix can produce this result during review. |
| 2 | The workbook or snapshot could not be read or validated, or the export destination conflicts with an existing snapshot. |

The snapshot contains each user-defined name, its workbook or worksheet scope, its saved formula text, its comment and its visibility. It groups names at the first period. An unprefixed name is checked as an alias when the first line of its name comment declares `Governed formula compatibility alias: ModuleName`. Worksheet cell formulas, table rows and calculation results are outside the snapshot. Use Excel to verify those dependencies and results.

## What the two checks catch

| Change in a separate example copy | Saved-definition checker | Recalculated Excel cases |
| --- | --- | --- |
| Delivered example | Matches the supplied snapshot | 8 of 8 pass |
| Change the second Beta to Theta | Definitions still match | 7 of 8 pass. Beta now resolves to ID-102. |
| Replace the function with a first-match lookup | Reports a change to `Lookup.ResolveID` | 6 of 8 pass. The ambiguous label and the reused ID both return an identifier. |

The first-match formula is a deliberate faulty replacement that demonstrates the tests catch the defect. The saved results and exact Excel build are in `verification.json`. To exercise the checker independently of Excel, run `python -B -m unittest test_formula_modules`.

## Compare version A with version B

Open the [worked approval example](approval/README.md) to compare the original workbook with four separate copies. A data-only rename earns a new decision after two focused checks. A faulty formula is rejected. Editing only the expected answer puts the workbook on hold for the rule owner. Editing both the caller and its expected answer makes the sheet show eight passes even though it returns the wrong ID for Beta.

The example includes downloadable workbooks, exact fingerprints, an independent set of eight agreed answers and a one-command replay. Each decision belongs to the bytes the reviewer examined.

## From review to release

Follow [the release owner’s three decisions](release/README.md). The owner makes B current after matching its fingerprint to the review. A stays current if C replaces B at the submission link or a mixed edit introduces a worksheet formula that still needs a test. Run `python -B release/demo.py` to try all three with the supplied workbooks.

The [proposed Power Automate flow](release/power-automate/README.md) shows how Riley’s approval in Dataverse identifies one SharePoint file version and SHA-256 fingerprint. Follow request **R-017** to see what Morgan checks before publishing. Run `python -B release/power-automate/verify_binding.py` to replay the example locally.

## Share the solution

Send the reviewed workbook when recipients need formulas they can run. To install the example in a different workbook, create the workbook-scoped name `Lookup.ResolveID` and paste the contents of [Lookup.ResolveID.txt](Lookup.ResolveID.txt) into Name Manager’s **Refers to** field. Pass in the destination’s label and ID columns. For other modules, also bring across their required names and data columns, then run the cases in that workbook. The JSON snapshot preserves Excel’s stored syntax for review.

Use this brief for the next change.

```text
Workbook version and Excel build:
Defined name and scope:
Business rule and expected return type or spill shape:
Inputs and dependencies:
Smallest failing input, expected result and current result:
Ordinary case that must keep working:
Affected callers and any intentional change to a counting population:
Candidate fingerprint, recalculated results and reviewer decision:
```

The broader proposal for governing operational workbooks is documented in [Power Platform guidance issue #110](https://github.com/microsoft/PowerPnPGuidanceHub/issues/110).
