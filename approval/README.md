# Which workbook does an approval cover?

Version A is the starting point: a label returns an ID only if the label matches one row and the ID is nonblank and unique. **Beta** matches two rows, with IDs **ID-102** and **ID-103**. Its agreed answer is empty text. Imagine the reviewer accepted [version A](../example.xlsx) with that rule and eight expected answers. Each version B below starts from a fresh copy of A.

In this example, a workbook decision applies to one SHA-256 fingerprint. Every B has different bytes and needs its own decision. A reviewer can reuse earlier evidence for an unchanged formula; a new fingerprint alone says nothing about whether the behavior is right. These are **illustrative decisions** grounded in actual Excel results. They do not represent sign-off by a person, Microsoft, or an approval service.

| Workbook and edit from A | Named formula | Sheet checks pass | Returned IDs match A | Expected cells match A | Example decision for these bytes |
| --- | --- | ---: | ---: | ---: | --- |
| [A: original workbook](../example.xlsx) | Matches | 8/8 | 8/8 | 8/8 | **Accept A** as the baseline. |
| [B: data only](B-data.xlsx) — rename the label for ID-108 | Matches | 8/8 | 8/8 | 8/8 | **Accept this B** after the data owner checks the rename. |
| [B: formula change](B-formula.xlsx) — use the first match | Changed | 6/8 | 6/8 | 8/8 | **Reject.** Two invalid selections now return IDs. |
| [B: expected answer change](B-expected.xlsx) — expect ID-102 for Beta | Matches | 7/8 | 8/8 | 7/8 | **Hold for the rule owner.** The result still meets the agreed rule. |
| [B: caller and expected answer change](B-combined.xlsx) — return and expect ID-102 for Beta | Matches | 8/8 | 7/8 | 7/8 | **Reject.** The sheet agrees with itself, but Beta resolves to the wrong ID. |

**How to read the counts:** “Sheet checks” is what column F displays. “Returned IDs” compares column E with the [eight expected answers recorded at A](https://github.com/wornpage/governed-excel-formula-module-kit/blob/535ba49f28f0c3087e0634e444d10d226782fb85/verification.json). “Expected cells” compares the answers entered in column D with that same external record. Empty text is a valid return; a number, including zero, is not. All three columns keep the same eight cases. The two rename checks below are counted separately.

## Try the review in Excel

1. Open A and one B in Excel with LAMBDA support. Work on copies if you plan to save. Recalculate, then compare **Cases** columns D (expected), E (returned) and F (check). The [decision record](decisions.json) lists each workbook’s full fingerprint, edits, failed rows, Excel build and example disposition.
2. **Data only:** Compare `Records!B13`: “Label A” became “Label B”; `C13` still contains ID-108. None of the eight cases selects either exact label. In a disposable copy of each workbook, enter “Label A” and then “Label B” in `Cases!C6`; read `E6` each time. A returns ID-108 then empty text. B returns empty text then ID-108. Discard the probe edits. The formula review carries forward; record a **new workbook decision** for B. A data change is not an automatic pass: renaming the second Beta instead [makes one of the eight cases fail](../README.md#try-it-in-excel).
3. **Formula change:** In Name Manager, compare `Lookup.ResolveID`. B takes the first match. `Cases!E7` returns ID-102 for Beta and `E11` returns ID-104 for Delta, whose ID another record shares. Both must be empty. Return B to the formula maintainer; A remains the usable version.
4. **Expected answer change:** Only `Cases!D7` changed, from blank to ID-102. `E7` still returns empty text; `F7` shows FAIL. Restore D7 or ask the rule owner to change the written rule. An intentional rule change needs a separately reviewed set of expected answers, an updated implementation, a fresh run and a new decision. An edit to the candidate’s own expected cell cannot authorize itself.
5. **Why an all-green sheet can still fail review:** In the last B, `Cases!E7` bypasses the named function and uses a direct first-match lookup. `D7` also says ID-102. `F7` now says PASS, and the named definition still matches. The reviewer rejects B because the independently recorded answer for Beta is empty. Restore both the caller and the expectation before retesting.

## Replay the supplied evidence

From the root of the extracted kit, run with Python 3.10 or later:

```text
python -B approval/replay.py
```

The replay checks the five fingerprints, compares saved names and worksheet edits, and reads the results Excel saved. It succeeds when it reproduces the table, **including the rejected examples**. It does not recalculate Excel, issue approvals, or assess other workbooks. For a new candidate, recalculate in the target Excel build, save and reopen a working copy, and record a new decision. Keep the approved workbook, Git revision, formula snapshot and expected answers independently of the candidate.

To run the general definition check alone, try `python formula_modules.py check approval/B-formula.xlsx` (changed name, exit 1) and `python formula_modules.py check approval/B-expected.xlsx` (matching name, exit 0). Neither exit code is a workbook approval. You can carry out this review locally without Dataverse or Power Automate.
