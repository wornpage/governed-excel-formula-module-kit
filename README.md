# Solving governed Excel formula modules

Start with a business rule someone else can test. Keep a readable copy of the named formulas and an Excel case that would fail if the rule were broken. This kit includes a fictional asset lookup you can open, change and check.

[Download the complete kit (ZIP)](https://github.com/wornpage/governed-excel-formula-module-kit/archive/refs/heads/main.zip)

## Try it in Excel

1. Open [example.xlsx](example.xlsx) in Excel with LAMBDA support. On **Cases**, all eight checks should say **PASS**. Two register records use the label **Pump North**. The expected key for that selection is empty because either key would identify just one of those records.
2. On a copy, change **Register!B8** from **Pump North** to **Pump South**. Recalculate. **Cases!E7** should now return **K-100**, and **F7** should say **FAIL**. The expected result still describes the original register. Undo the edit and the check should return to **PASS**.
3. On Windows, open **Formulas → Name Manager** and inspect `Asset.ResolveKey`. The name is scoped to the workbook. Its comment explains the inputs and return value. [Microsoft’s LAMBDA instructions](https://support.microsoft.com/en-us/excel/functions/lambda-function)

The function takes a selected label and two aligned, one-column ranges of labels and text keys. It returns a single text key only when both the label and its nonblank key occur once. Otherwise it returns empty text. Matching ignores capitalization and treats an asterisk as a literal character. The function looks for the first and last exact match. Different positions mean the value occurs more than once. [Microsoft’s XMATCH reference](https://support.microsoft.com/en-us/excel/functions/xmatch-function)

## Apply the method to your workbook

| Step | What the reviewer needs |
| --- | --- |
| 1. Define the rule | Record the full name and scope, inputs and units, expected value or spill dimensions, and what blanks, zeros, errors and duplicate keys mean. For a KPI, state its numerator, denominator, filters and period. |
| 2. Capture the baseline | Save a copy of the workbook. Export its named definitions into a new directory. The manifest identifies the workbook by its SHA-256 fingerprint. Keep this directory as the comparison point. |
| 3. Reproduce the failure | Use the smallest inputs that trigger the defect. Write the expected answer from the rule before changing the formula. Run the original formula and record its actual answer. Keep at least one ordinary case that already works. |
| 4. Fix the owning name | Follow the failing worksheet cell to the named formula and its dependencies. Change the name that owns the rule. Review affected names, aliases, worksheet formulas, validation rules and other consumers. |
| 5. Verify and hand off | Compare the candidate against the baseline. In the target Excel version, recalculate, save and reopen a working copy. Re-run the failing case and affected callers. Record expected and actual results, Excel build, workbook fingerprint and the reviewer’s decision. Preserve the prior workbook and snapshot for rollback. |

For an intentional policy change, write the new expectation and its reason alongside the old one. For example, the share of assets with an open action divides distinct in-scope assets with at least one open action by all distinct in-scope assets. One asset with three open actions contributes one to the numerator. Record any change to the eligible population.

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
| Rename the second Pump North to Pump South | Definitions still match | 7 of 8 pass. Pump North now resolves to K-100. |
| Replace the function with a first-match lookup | Reports a change to `Asset.ResolveKey` | 6 of 8 pass. The ambiguous label and the reused key both return an identifier. |

The first-match formula is a deliberate faulty replacement that demonstrates the tests catch the defect. The saved results and exact Excel build are in `verification.json`. To exercise the checker independently of Excel, run `python -B -m unittest test_formula_modules`.

## Share the solution

Send the reviewed workbook when recipients need formulas they can run. To install the example in a different workbook, create the workbook-scoped name `Asset.ResolveKey` and paste the contents of [Asset.ResolveKey.txt](Asset.ResolveKey.txt) into Name Manager’s **Refers to** field. Pass in the destination’s label and key columns. For other modules, also bring across their required names and data columns, then run the cases in that workbook. The JSON snapshot preserves Excel’s stored syntax for review.

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
