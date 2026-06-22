"""
ci_validate.py — CI/CD validation script for boAt Review Analyzer
Runs all 6 review pipelines (real API calls, same as the Monday job) but
does NOT send the consolidated email — instead it inspects each generated
Excel file and verifies the structure is correct, then prints a pass/fail
report and exits non-zero if anything looks wrong.

This is meant to run automatically on every push / pull request via
.github/workflows/ci.yml, so logic bugs (wrong column numbers, missing
sheets, blank required fields) are caught immediately instead of only
being discovered after a bad weekly report goes out.
"""

import os
import sys
import openpyxl

# Reuse the exact same PIPELINES list and run_pipeline() function from
# run_all.py — this guarantees CI tests the real code path, not a
# reimplementation that could drift out of sync.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run_all import PIPELINES, run_pipeline, BASE_DIR

# ── EXPECTED STRUCTURE PER APP ────────────────────────────────────────────────
# Each entry: required sheet names, and required non-empty columns in the
# main data sheet (by header name) that must have at least SOME populated
# values across all rows if there is at least one row of data.
EXPECTED = {
    "Android - CrestApp": {
        "main_sheet": "CrestApp_Review_Data",
        "required_sheets": ["CrestApp_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User Name", "Ratings", "Reviews", "Date Time"],
    },
    "Android - HearablesApp": {
        "main_sheet": "Happ_Review_Data",
        "required_sheets": ["Happ_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User Name", "Ratings", "Reviews", "Date Time"],
    },
    "Android - WearablesApp": {
        "main_sheet": "WearablesApp_Review_Data",
        "required_sheets": ["WearablesApp_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User Name", "Ratings", "Reviews", "Date Time"],
    },
    "iOS - CrestApp": {
        "main_sheet": "CrestApp_Review_Data",
        "required_sheets": ["CrestApp_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User", "Rating", "Review", "Date"],
    },
    "iOS - HearablesApp": {
        "main_sheet": "Happ_Review_Data",
        "required_sheets": ["Happ_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User", "Rating", "Review", "Date"],
    },
    "iOS - WearablesApp": {
        "main_sheet": "Wapp_Review_Data",
        "required_sheets": ["Wapp_Review_Data", "Positive_Reviews", "Negative_Reviews", "Summary"],
        "required_columns": ["reviewId", "User", "Rating", "Review", "Date"],
    },
}

# ── VALIDATE ONE EXCEL FILE ───────────────────────────────────────────────────
def validate_xlsx(label, xlsx_path):
    spec = EXPECTED[label]
    errors = []

    if not os.path.exists(xlsx_path):
        # No reviews this week is an acceptable outcome, not a failure —
        # we already log it separately, just skip structural checks.
        return ["NO_DATA"]

    try:
        wb = openpyxl.load_workbook(xlsx_path)
    except Exception as e:
        return [f"Could not open workbook: {e}"]

    for sheet_name in spec["required_sheets"]:
        if sheet_name not in wb.sheetnames:
            errors.append(f"Missing expected sheet: '{sheet_name}'")

    if spec["main_sheet"] in wb.sheetnames:
        ws = wb[spec["main_sheet"]]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

        for col_name in spec["required_columns"]:
            if col_name not in headers:
                errors.append(f"Missing expected column '{col_name}' in '{spec['main_sheet']}'")
                continue
            col_idx = headers.index(col_name) + 1
            non_empty = any(
                ws.cell(r, col_idx).value not in (None, "")
                for r in range(2, ws.max_row + 1)
            )
            if ws.max_row > 1 and not non_empty:
                errors.append(f"Column '{col_name}' is completely empty in '{spec['main_sheet']}'")

        # Sanity check: App Version column should never contain an OS-version
        # looking value like a bare single/double digit SDK number — this is
        # exactly the bug we hit before (column 6 vs 9 mixup).
        if "App Version" in headers:
            app_ver_idx = headers.index("App Version") + 1
            for r in range(2, min(ws.max_row + 1, 12)):  # spot-check first 10 rows
                val = ws.cell(r, app_ver_idx).value
                if val and str(val).strip().isdigit() and len(str(val).strip()) <= 2:
                    errors.append(
                        f"Suspicious 'App Version' value '{val}' in row {r} — "
                        f"looks like an OS SDK number, not an app version (e.g. 5.4.0)"
                    )
                    break

    if "Summary" in wb.sheetnames:
        ws_sum = wb["Summary"]
        review_date_cell = None
        for col in range(1, ws_sum.max_column + 1):
            if ws_sum.cell(1, col).value in ("Review Data", "Review Date"):
                review_date_cell = ws_sum.cell(2, col).value
                break
        if review_date_cell and "None" in str(review_date_cell):
            errors.append(f"Summary date range looks broken: '{review_date_cell}'")

    wb.close()
    return errors

# ── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("#" * 70)
    print("# CI VALIDATION — running all 6 pipelines (no email will be sent)")
    print("#" * 70)

    all_errors = {}
    for pipeline in PIPELINES:
        xlsx_path = run_pipeline(pipeline)
        label = pipeline["label"]
        errors = validate_xlsx(label, xlsx_path) if xlsx_path else ["NO_DATA"]
        all_errors[label] = errors

    print("\n" + "#" * 70)
    print("# VALIDATION RESULTS")
    print("#" * 70)

    has_real_failure = False
    for label, errors in all_errors.items():
        if errors == ["NO_DATA"]:
            print(f"  ⚠️  SKIPPED (no reviews this week) — {label}")
        elif not errors:
            print(f"  ✅ PASS — {label}")
        else:
            has_real_failure = True
            print(f"  ❌ FAIL — {label}")
            for e in errors:
                print(f"       - {e}")

    print("\n" + "#" * 70)
    if has_real_failure:
        print("# CI FAILED — see errors above")
        print("#" * 70)
        sys.exit(1)
    else:
        print("# CI PASSED")
        print("#" * 70)
        sys.exit(0)
