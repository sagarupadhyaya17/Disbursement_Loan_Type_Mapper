import pandas as pd
import numpy as np

# ==== File Paths ====
disbursement_file = r"Input_Files/Sheet1.xlsx"
ytd_file = r"Input_Files/8.Disbursement Fagun 2082.xlsx"
main_file = r"Input_Files/Duelist 19th Feb, 2026.xlsx"
output_file = r"Output_Files/updated_disbursement.xlsx"

print("Processing... Please wait ⏳")

# ==== Read Excel Files ====
disb_df = pd.read_excel(disbursement_file, dtype=str)
ytd_df = pd.read_excel(ytd_file, sheet_name='YTD', dtype=str)
main_df = pd.read_excel(main_file, sheet_name='Mainsheet', dtype=str)

# =====================================================
# ✅ STEP 1 — CLEAN COLUMN NAMES
# =====================================================
def clean_columns(df):
    df.columns = df.columns.str.strip()
    return df.astype(str)

disb_df = clean_columns(disb_df)
ytd_df = clean_columns(ytd_df)
main_df = clean_columns(main_df)

# =====================================================
# ✅ STEP 2 — DYNAMIC COLUMN RENAME
# =====================================================
# Rename ACType
for col in disb_df.columns:
    if col.replace(" ", "").lower() == "actype":
        disb_df.rename(columns={col: "AcType"}, inplace=True)

for col in ytd_df.columns:
    if col.replace(" ", "").lower() == "actype":
        ytd_df.rename(columns={col: "AcType"}, inplace=True)

for col in main_df.columns:
    if col.replace(" ", "").lower() in ["at", "actype"]:
        main_df.rename(columns={col: "AcType"}, inplace=True)

# Rename Loan Type
for col in disb_df.columns:
    if col.replace(" ", "").lower() in ["oldacnum", "loantype"]:
        disb_df.rename(columns={col: "Loan Type"}, inplace=True)

for col in ytd_df.columns:
    if col.replace(" ", "").lower() in ["loantype", "loantype"]:
        ytd_df.rename(columns={col: "Loan Type"}, inplace=True)

for col in main_df.columns:
    if col.replace(" ", "").lower() == "loantype":
        main_df.rename(columns={col: "Loan Type"}, inplace=True)

# =====================================================
# ✅ STEP 3 — CLEAN KEY COLUMN VALUES
# =====================================================
for col in ["AcType", "Loan Type", "BranchName"]:
    if col in ytd_df.columns:
        ytd_df[col] = ytd_df[col].astype(str).str.strip()

    if col in disb_df.columns:
        disb_df[col] = disb_df[col].astype(str).str.strip()

    if col in main_df.columns:
        main_df[col] = main_df[col].astype(str).str.strip()

# Convert empty loan type to NaN
disb_df["Loan Type"] = disb_df["Loan Type"].replace(["", "nan", "None"], pd.NA)

disb_df = disb_df[(disb_df['AcType']!='4Z')]

# =====================================================
# ✅ STEP 4 — FUNCTION TO FETCH LOAN TYPE
# =====================================================
def get_loan_type(row, ref_df):
    actype = row["AcType"]
    branch = row["BranchName"]

    temp = ref_df[ref_df["AcType"] == actype]

    if temp.empty:
        return np.nan

    unique_loans = temp["Loan Type"].dropna().unique()

    # Case 1 → single loan type
    if len(unique_loans) == 1:
        return unique_loans[0]

    # Case 2 → multiple loan type → filter by branch
    temp2 = temp[temp["BranchName"] == branch]
    unique_loans2 = temp2["Loan Type"].dropna().unique()

    if len(unique_loans2) == 1:
        return unique_loans2[0]

    return np.nan

# =====================================================
# ✅ STEP 5 — APPLY LOGIC
# =====================================================
# First search in MAIN
disb_df["Loan Type"] = disb_df.apply(
    lambda row: get_loan_type(row, main_df) if pd.isna(row["Loan Type"]) else row["Loan Type"],
    axis=1
)

# Then fallback to YTD
disb_df["Loan Type"] = disb_df.apply(
    lambda row: get_loan_type(row, ytd_df) if pd.isna(row["Loan Type"]) else row["Loan Type"],
    axis=1
)

# =====================================================
# ✅ STEP 6 — EXPORT RESULT
# =====================================================
disb_df.to_excel(output_file, index=False)

print("✅ Loan Type mapping completed!")
print("📁 Output saved to:", output_file)

# =====================================================
# ✅ STEP 7 — SHOW UNMATCHED (optional)
# =====================================================
unmatched = disb_df[disb_df["Loan Type"].isna()]
print("⚠ Unmatched rows:", len(unmatched))

input("Press Enter to exit...")