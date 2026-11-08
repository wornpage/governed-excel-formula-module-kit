"""Replay the proposed approval handoff with local workbooks and fictional tenant records."""
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from demo import Decision, RELEASE_OWNER, checked_inputs, digest, publish, require

RILEY_OBJECT_ID = "b1100000-0000-4000-8000-000000000001"


def version_key(file_version):
    return tuple(file_version[field] for field in ("driveId", "itemId", "versionId"))


def release_result(review, approval, response, observed_version, staged_file, current, verified_sha):
    """Join one eligible review, its native reply and the observed SharePoint version."""
    if (not review["approvalId"] or review["approvalId"] != approval["id"]
            or review["assignedReviewerObjectId"] != RILEY_OBJECT_ID
            or approval["assignedReviewerObjectId"] != RILEY_OBJECT_ID
            or review["reviewedSha256"] != verified_sha):
        return "HOLD"
    if (response is None or response["approvalId"] != approval["id"]
            or response["responderObjectId"] != RILEY_OBJECT_ID or not response["respondedAtUtc"]):
        return "HOLD"
    if response["outcome"] == "Reject":
        return "REJECT"
    if response["outcome"] != "Approve":
        return "HOLD"
    expected_key = version_key(review["sharePointVersion"])
    if not all(expected_key) or version_key(observed_version) != expected_key:
        return "STALE"
    # The release demo reads the staged bytes, verifies the full digest and publishes those bytes.
    return publish(RELEASE_OWNER, Decision("ACCEPT", verified_sha), staged_file, current)


def main():
    sample = json.loads((HERE / "example.json").read_text(encoding="utf-8"))
    a, b, c, *_ = checked_inputs()
    review = sample["workbookReview"]
    approval, reply = sample["nativeApproval"], sample["nativeResponse"]
    versions = {b.name: b, c.name: c}
    expected_evidence = {
        "gitRevision": "8a69a6cbd4dee4ff57e375c9932252c66e79b05a",
        "file": "approval/decisions.json",
        "version": "B-data",
    }
    require(review["evidence"] == expected_evidence, "Unknown or missing test evidence for R-017.")
    verified_sha = digest(b.read_bytes())  # checked_inputs has verified B's Excel receipt and rename probes.
    require(review["reviewedSha256"] == verified_sha, "The protected request no longer identifies tested B.")
    print(f"{review['requestId']} | approval {approval['id']} | reviewed SHA-256 {verified_sha}")
    with TemporaryDirectory() as folder:
        current = Path(folder) / "Current.xlsx"
        staged = Path(folder) / "Staged.xlsx"

        def run(label, observed, source, response, expected, expected_current):
            current.write_bytes(a.read_bytes())  # Independent scenarios, each starts at A.
            staged.write_bytes(source.read_bytes())
            actual = release_result(review, approval, response, observed, staged, current, verified_sha)
            require(actual == expected and current.read_bytes() == expected_current.read_bytes(),
                    f"{label}: unexpected decision or published bytes.")
            letter = "A" if expected_current == a else "B"
            print(f"{actual:<8} | {label} | observed {observed['itemId']} v{observed['versionId']} "
                  f"{digest(staged.read_bytes())[:12]} | current {letter}")

        for case in sample["publicationChecks"]:
            workbook = versions[case["workbook"]]
            require(digest(workbook.read_bytes()) == case["observedSha256"], "Unexpected observed workbook.")
            run(case["label"], case["currentSubmission"], workbook, reply, case["expectedResult"],
                {"A": a, "B": b}[case["expectedCurrent"]])

        matched = review["sharePointVersion"]
        run("Same version label, different bytes", matched, c, reply, "STALE", a)
        run("Same bytes, different file item", {**matched, "itemId": "workbook-copy"}, b, reply, "STALE", a)
        run("Reply belongs to another approval", matched, b, {**reply, "approvalId": "other"}, "HOLD", a)
        run("Reply from another reviewer", matched, b, {**reply, "responderObjectId": "other"}, "HOLD", a)
        run("Timeout without a response", matched, b, None, "HOLD", a)
        run("Riley rejects B", matched, b, {**reply, "outcome": "Reject"}, "REJECT", a)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError) as exc:
        sys.exit(f"Cannot replay the proposed flow: {exc}")
