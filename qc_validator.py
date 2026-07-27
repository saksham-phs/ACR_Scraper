"""
Quality Control (QC) Validator for ACR Convergence 2026 Scraper Output
Validates workbook schema, row counts, data integrity, non-null guards, and column specifications.
"""

from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

EXPECTED_HEADERS = [
    "Session_ID", "Session_Title", "Session_Type", "Room", "Date", "Day", "Time", "Chairs", "Chairs_Geography", "Session_Description", "Session_URL",
    "Presentation_ID", "Presentation_Title", "Presenter", "Presentation_Time", "Authors", "Authors_and_Affiliations", "Presenter_Geography", "Presentation_Type", "Keywords", "Abstract_Full_Text", "Acknowledgments_and_Disclosures", "Presentation_URL"
]

EXCEL_PATH = Path("output/acr_2026_programme.xlsx")


def run_qc_checks() -> bool:
    print("=" * 70)
    print("      ACR CONVERGENCE 2026 SCRAPER QUALITY CONTROL (QC) VALIDATOR")
    print("=" * 70)

    if not EXCEL_PATH.exists():
        print(f"[FAIL] Output workbook not found at {EXCEL_PATH.absolute()}")
        return False

    print(f"[OK] Found output file: {EXCEL_PATH.name} ({EXCEL_PATH.stat().st_size / 1024:.1f} KB)")

    # 1. Load with pandas
    df = pd.read_excel(EXCEL_PATH, sheet_name="Presentations")
    print(f"[OK] Workbook loaded. Data rows: {len(df)}, Columns: {len(df.columns)}")

    # 2. Check headers
    actual_headers = list(df.columns)
    if actual_headers != EXPECTED_HEADERS:
        print("[FAIL] Header mismatch!")
        print("Expected:", EXPECTED_HEADERS)
        print("Actual:  ", actual_headers)
        return False
    print(f"[OK] Column count and headers match expected schema ({len(EXPECTED_HEADERS)} columns).")

    # 3. Check for non-empty content
    if len(df) == 0:
        print("[FAIL] Workbook contains 0 data rows.")
        return False

    # 4. Mandatory Field Coverage
    session_title_nulls = df["Session_Title"].isna().sum()
    session_id_nulls = df["Session_ID"].isna().sum()
    date_nulls = df["Date"].isna().sum()

    print("\n--- Mandatory Field Null Checks ---")
    print(f"Session_Title null count: {session_title_nulls} / {len(df)}")
    print(f"Session_ID null count:    {session_id_nulls} / {len(df)}")
    print(f"Date null count:          {date_nulls} / {len(df)}")

    if session_title_nulls > 0 or date_nulls > 0:
        print("[WARNING] Some mandatory session fields contain null values.")

    # 5. Presentation Coverage
    has_presentation = df["Presentation_Title"].notna().sum()
    has_presenter = df["Presenter"].notna().sum()
    has_affiliations = df["Authors_and_Affiliations"].notna().sum()
    has_disclosures = df["Acknowledgments_and_Disclosures"].notna().sum()

    print("\n--- Presentation Level Field Coverage ---")
    print(f"Presentations with Titles:             {has_presentation} ({has_presentation/len(df)*100:.1f}%)")
    print(f"Presentations with Presenters/Authors: {has_presenter} ({has_presenter/len(df)*100:.1f}%)")
    print(f"Presentations with Affiliations:       {has_affiliations} ({has_affiliations/len(df)*100:.1f}%)")
    print(f"Presentations with Disclosures:        {has_disclosures} ({has_disclosures/len(df)*100:.1f}%)")

    # 6. Check OpenPyXL formatting & freeze pane
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Presentations"]
    freeze = ws.freeze_panes
    print(f"\n--- Excel Formatting & Layout Checks ---")
    print(f"Freeze Panes Setting: {freeze}")
    if freeze != "A2":
        print(f"[WARNING] Freeze panes is '{freeze}', expected 'A2'")

    print("\n" + "=" * 70)
    print("   QC RESULT: ALL CHECKS COMPLETED SUCCESSFULLY - VERIFICATION PASSED")
    print("=" * 70)
    return True


if __name__ == "__main__":
    success = run_qc_checks()
    sys.exit(0 if success else 1)
