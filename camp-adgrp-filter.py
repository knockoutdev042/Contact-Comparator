import pandas as pd
import numpy as np
from pathlib import Path

# ======================================================
# CONFIGURATION
# ======================================================

# Input file (.csv / .xlsx / .xls)
INPUT_FILE = "input.csv"

# Output Excel file
OUTPUT_FILE = "filtered_output8.xlsx"

# Column names
CAMPAIGN_ID_COL = "Campaign Id"
ADGROUP_ID_COL = "AdGroup Id"

# ======================================================
# HELPER FUNCTION
# ======================================================

def normalize_id(value):
    """
    Normalize IDs by:
    - Removing scientific notation
    - Removing decimals
    - Converting to proper integer string

    Example:
    2.02307E+11 -> 202307000000
    """

    if pd.isna(value):
        return np.nan

    value = str(value).strip()

    # Handle empty values
    if value.lower() in [
        "",
        "nan",
        "none",
        "null"
    ]:
        return np.nan

    try:
        return str(int(float(value)))
    except:
        return value


# ======================================================
# READ INPUT FILE
# ======================================================

print("Reading input file...")

file_extension = Path(INPUT_FILE).suffix.lower()

if file_extension == ".csv":

    df = pd.read_csv(
        INPUT_FILE,
        dtype=str
    )

elif file_extension in [".xlsx", ".xls"]:

    df = pd.read_excel(
        INPUT_FILE,
        dtype=str,
        engine="openpyxl"
    )

else:
    raise ValueError(
        "Unsupported file format. Use CSV or Excel."
    )

print("File loaded successfully.")

# ======================================================
# TOTAL ROW COUNT
# ======================================================

original_row_count = len(df)

print(f"Total Rows Found: {original_row_count}")

# ======================================================
# NORMALIZE IDS
# ======================================================

print("Normalizing Campaign and AdGroup IDs...")

df[CAMPAIGN_ID_COL] = df[CAMPAIGN_ID_COL].apply(normalize_id)
df[ADGROUP_ID_COL] = df[ADGROUP_ID_COL].apply(normalize_id)

# ======================================================
# CLEAN EMPTY VALUES
# ======================================================

empty_values = [
    "",
    " ",
    "nan",
    "NaN",
    "null",
    "NULL",
    "none",
    "None"
]

df[CAMPAIGN_ID_COL] = (
    df[CAMPAIGN_ID_COL]
    .astype(str)
    .str.strip()
    .replace(empty_values, np.nan)
)

df[ADGROUP_ID_COL] = (
    df[ADGROUP_ID_COL]
    .astype(str)
    .str.strip()
    .replace(empty_values, np.nan)
)

# ======================================================
# CONDITION 1
# BOTH IDs PRESENT
# ======================================================

condition_both = (
    df[CAMPAIGN_ID_COL].notna()
    &
    df[ADGROUP_ID_COL].notna()
)

df_both = df[condition_both].copy()

# ======================================================
# CONDITION 2
# ONLY ONE ID PRESENT
# ======================================================
# campaignid exists and adgroupid missing
# OR
# adgroupid exists and campaignid missing

condition_only_one = (

    (
        df[CAMPAIGN_ID_COL].notna()
        &
        df[ADGROUP_ID_COL].isna()
    )

    |

    (
        df[CAMPAIGN_ID_COL].isna()
        &
        df[ADGROUP_ID_COL].notna()
    )
)

df_only_one = df[condition_only_one].copy()

# ======================================================
# FINAL OUTPUT
# ======================================================
# Merge:
# - BOTH IDs present
# - ONLY ONE ID present

final_df = pd.concat(
    [df_both, df_only_one],
    ignore_index=True
)

# ======================================================
# REMOVE DUPLICATES (OPTIONAL SAFETY)
# ======================================================

final_df = final_df.drop_duplicates()

# ======================================================
# COUNTS
# ======================================================

both_count = len(df_both)

only_one_count = len(df_only_one)

final_count = len(final_df)

removed_count = (
    original_row_count - final_count
)

# ======================================================
# SUMMARY
# ======================================================

summary_df = pd.DataFrame({

    "Metric": [

        "Original Total Rows",

        "Rows with BOTH IDs",

        "Rows with ONLY ONE ID",

        "Rows Removed (Both Empty)",

        "Final Output Rows"
    ],

    "Count": [

        original_row_count,

        both_count,

        only_one_count,

        removed_count,

        final_count
    ]
})

print("\n========== SUMMARY ==========")

print(summary_df)

# ======================================================
# EXPORT OUTPUT
# ======================================================

print("\nGenerating output Excel file...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    # Final merged output
    final_df.to_excel(
        writer,
        sheet_name="Filtered_Data",
        index=False
    )

    # Both IDs present
    df_both.to_excel(
        writer,
        sheet_name="Both_IDs_Present",
        index=False
    )

    # Only one ID present
    df_only_one.to_excel(
        writer,
        sheet_name="Only_One_ID_Present",
        index=False
    )

    # Summary
    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

print("\n===================================")
print("Output file generated successfully!")
print(f"Output File: {OUTPUT_FILE}")
print("===================================")