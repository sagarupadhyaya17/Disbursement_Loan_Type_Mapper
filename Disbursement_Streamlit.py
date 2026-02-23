import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Loan Type Mapper", layout="wide")

st.title("🏦 Loan Type Mapping Tool")
st.write("Upload required files and generate mapped disbursement report.")

# =====================================================
# FILE UPLOAD SECTION
# =====================================================
col1, col2, col3 = st.columns(3)

with col1:
    disbursement_file = st.file_uploader("Upload Disbursement File", type=["xlsx"])

with col2:
    ytd_file = st.file_uploader("Upload YTD File", type=["xlsx"])

with col3:
    main_file = st.file_uploader("Upload Main File", type=["xlsx"])


# =====================================================
# HELPER FUNCTION
# =====================================================
def prepare_df(df):
    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Remove duplicate columns
    df = df.loc[:, ~df.columns.duplicated()]

    return df


# =====================================================
# PROCESS BUTTON
# =====================================================
if st.button("🚀 Run Mapping"):

    if not (disbursement_file and ytd_file and main_file):
        st.error("Please upload all three files.")
        st.stop()

    with st.spinner("Processing mapping..."):

        # Read files
        disb_df = pd.read_excel(disbursement_file)
        ytd_df = pd.read_excel(ytd_file, sheet_name="YTD")
        main_df = pd.read_excel(main_file, sheet_name="Mainsheet")

        # Clean columns
        disb_df = prepare_df(disb_df)
        ytd_df = prepare_df(ytd_df)
        main_df = prepare_df(main_df)

        # =====================================================
        # STANDARDIZE COLUMN NAMES
        # =====================================================
        def standardize(df):
            rename_dict = {}

            for col in df.columns:
                key = col.replace(" ", "").lower()

                if key in ["actype", "at"]:
                    rename_dict[col] = "AcType"

                if key in ["loantype", "oldacnum"]:
                    rename_dict[col] = "Loan Type"

                if key in ["branchname", "branch"]:
                    rename_dict[col] = "BranchName"

            df = df.rename(columns=rename_dict)

            # Remove duplicates after renaming
            df = df.loc[:, ~df.columns.duplicated()]

            return df

        disb_df = standardize(disb_df)
        ytd_df = standardize(ytd_df)
        main_df = standardize(main_df)

        # =====================================================
        # CLEAN VALUES
        # =====================================================
        for df in [disb_df, ytd_df, main_df]:
            for col in ["AcType", "Loan Type", "BranchName"]:
                if col in df.columns:
                    df[col] = (
                        df[col]
                        .astype(str)
                        .str.strip()
                        .replace(["", "nan", "None"], np.nan)
                    )

        # Remove AcType 4Z
        if "AcType" in disb_df.columns:
            disb_df = disb_df[disb_df["AcType"] != "4Z"]

        # =====================================================
        # MAPPING FUNCTION
        # =====================================================
        def get_loan_type(row, ref_df):

            if "AcType" not in ref_df.columns or "Loan Type" not in ref_df.columns:
                return np.nan

            actype = row.get("AcType")
            branch = row.get("BranchName")

            temp = ref_df[ref_df["AcType"] == actype]

            if temp.empty:
                return np.nan

            unique_loans = temp["Loan Type"].dropna().unique()

            if len(unique_loans) == 1:
                return unique_loans[0]

            if branch and "BranchName" in ref_df.columns:
                temp2 = temp[temp["BranchName"] == branch]
                unique_loans2 = temp2["Loan Type"].dropna().unique()

                if len(unique_loans2) == 1:
                    return unique_loans2[0]

            return np.nan

        # =====================================================
        # APPLY LOGIC
        # =====================================================
        disb_df["Loan Type"] = disb_df.apply(
            lambda row: get_loan_type(row, main_df)
            if pd.isna(row.get("Loan Type"))
            else row["Loan Type"],
            axis=1
        )

        disb_df["Loan Type"] = disb_df.apply(
            lambda row: get_loan_type(row, ytd_df)
            if pd.isna(row.get("Loan Type"))
            else row["Loan Type"],
            axis=1
        )

        # =====================================================
        # SHOW DATA
        # =====================================================
        tab1, tab2 = st.tabs(["📊 Mapped Data", "⚠ Unmatched"])

        with tab1:
            st.dataframe(disb_df, use_container_width=True)

        unmatched = disb_df[disb_df["Loan Type"].isna()]

        with tab2:
            st.write(f"Unmatched rows: {len(unmatched)}")
            st.dataframe(unmatched, use_container_width=True)

        # =====================================================
        # DOWNLOAD
        # =====================================================
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            disb_df.to_excel(writer, index=False, sheet_name="Mapped_Data")

        buffer.seek(0)

        st.download_button(
            label="📥 Download Result",
            data=buffer,
            file_name="updated_disbursement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ Mapping Completed!")