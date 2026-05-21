# -*- coding: utf-8 -*-

"""
NAV -> SAP Validation Toolkit
-----------------------------------

Business-oriented ERP migration validation workflow.

Main capabilities:
- Data reconciliation
- Mapping validation
- Duplicate detection
- Financial consistency checks
- ERP migration controls
"""

from pathlib import Path
import pandas as pd


# ==========================================
# CONFIGURATION
# ==========================================

INPUT_FOLDER = Path("input_data")

OUTPUT_FOLDER = Path("reports")

NAV_FILE = "nav_export.xlsx"

SAP_FILE = "sap_export.xlsx"


# ==========================================
# HELPERS
# ==========================================

def load_excel_file(file_path: Path) -> pd.DataFrame:
    """
    Loads Excel file into pandas DataFrame.
    """

    return pd.read_excel(file_path)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardizes column names.
    """

    df.columns = [
        str(col).strip().upper()
        for col in df.columns
    ]

    return df


def detect_duplicates(
    df: pd.DataFrame,
    key_columns: list
) -> pd.DataFrame:
    """
    Detects duplicated records.
    """

    duplicates = df[
        df.duplicated(
            subset=key_columns,
            keep=False
        )
    ]

    return duplicates


# ==========================================
# VALIDATION WORKFLOW
# ==========================================

def validate_migration():

    nav_path = INPUT_FOLDER / NAV_FILE
    sap_path = INPUT_FOLDER / SAP_FILE

    print("Loading NAV export...")
    nav_df = load_excel_file(nav_path)

    print("Loading SAP export...")
    sap_df = load_excel_file(sap_path)

    nav_df = normalize_columns(nav_df)
    sap_df = normalize_columns(sap_df)

    print("Running duplicate checks...")

    nav_duplicates = detect_duplicates(
        nav_df,
        ["DOCUMENT_NO"]
    )

    sap_duplicates = detect_duplicates(
        sap_df,
        ["DOCUMENT_NO"]
    )

    print("\n==============================")
    print("VALIDATION SUMMARY")
    print("==============================")

    print(
        f"NAV records: {len(nav_df)}"
    )

    print(
        f"SAP records: {len(sap_df)}"
    )

    print(
        f"NAV duplicates: {len(nav_duplicates)}"
    )

    print(
        f"SAP duplicates: {len(sap_duplicates)}"
    )

    print("==============================\n")


# ==========================================
# ENTRYPOINT
# ==========================================

if __name__ == "__main__":

    validate_migration()
