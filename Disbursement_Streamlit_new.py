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
    disbursement_file = st.file_uploader(
        "Upload Disbursement File",
        type=["xlsx", "xlsb"]
    )

with col2:
    ytd_file = st.file_uploader(
        "Upload YTD File",
        type=["xlsx", "xlsb"]
    )

with col3:
    main_file = st.file_uploader(
        "Upload Duelist File",
        type=["xlsx", "xlsb"]
    )

# =====================================================
# HELPER FUNCTION
# =====================================================
def prepare_df(df):

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Remove duplicate columns
    df = df.loc[:, ~df.columns.duplicated()]

    return df.astype(str)

# =====================================================
# STANDARDIZE COLUMN NAMES
# =====================================================
def standardize(df):

    rename_dict = {}

    for col in df.columns:

        key = col.replace(" ", "").lower()

        # AcType
        if key in ["actype", "at"]:
            rename_dict[col] = "AcType"

        # Loan Type
        if key in ["loantype", "oldacnum"]:
            rename_dict[col] = "Loan Type"

        # Branch Name
        if key in ["branchname", "branch"]:
            rename_dict[col] = "BranchName"

    df = df.rename(columns=rename_dict)

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

        # =====================================================
        # READ DISBURSEMENT FILE
        # =====================================================
        if disbursement_file.name.endswith(".xlsb"):

            disb_df = pd.read_excel(
                disbursement_file,
                dtype=str,
                engine="pyxlsb"
            )

        else:

            disb_df = pd.read_excel(
                disbursement_file,
                dtype=str
            )

        # =====================================================
        # READ YTD FILE
        # =====================================================
        if ytd_file.name.endswith(".xlsb"):

            ytd_df = pd.read_excel(
                ytd_file,
                sheet_name="YTD",
                dtype=str,
                engine="pyxlsb"
            )

        else:

            ytd_df = pd.read_excel(
                ytd_file,
                sheet_name="YTD",
                dtype=str
            )

        # =====================================================
        # READ DUELIST FILE
        # =====================================================
        if main_file.name.endswith(".xlsb"):

            main_df = pd.read_excel(
                main_file,
                sheet_name="Mainsheet",
                dtype=str,
                engine="pyxlsb"
            )

        else:

            main_df = pd.read_excel(
                main_file,
                sheet_name="Mainsheet",
                dtype=str
            )

        # =====================================================
        # CLEAN DATAFRAMES
        # =====================================================
        disb_df = prepare_df(disb_df)
        ytd_df = prepare_df(ytd_df)
        main_df = prepare_df(main_df)

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

        # =====================================================
        # REMOVE 4Z
        # =====================================================
        if "AcType" in disb_df.columns:

            disb_df = disb_df[
                disb_df["AcType"] != "4Z"
            ]

        # =====================================================
        # CREATE DUELIST MAPPING
        # =====================================================
        main_map = main_df[
            main_df["Loan Type"].notna()
        ][["AcType", "BranchName", "Loan Type"]].drop_duplicates()

        # =====================================================
        # FIRST MERGE → DUELIST
        # =====================================================
        disb_df = disb_df.merge(
            main_map,
            on=["AcType", "BranchName"],
            how="left",
            suffixes=("", "_main")
        )

        # Fill blank Loan Type from Duelist
        disb_df["Loan Type"] = disb_df["Loan Type"].fillna(
            disb_df["Loan Type_main"]
        )

        # Remove extra column
        disb_df.drop(
            columns=["Loan Type_main"],
            inplace=True
        )

        # =====================================================
        # CREATE YTD MAPPING
        # =====================================================
        ytd_map = ytd_df[
            ytd_df["Loan Type"].notna()
        ][["AcType", "BranchName", "Loan Type"]].drop_duplicates()

        # =====================================================
        # SECOND MERGE → YTD
        # =====================================================
        disb_df = disb_df.merge(
            ytd_map,
            on=["AcType", "BranchName"],
            how="left",
            suffixes=("", "_ytd")
        )

        # Fill remaining blank Loan Type
        disb_df["Loan Type"] = disb_df["Loan Type"].fillna(
            disb_df["Loan Type_ytd"]
        )

        # Remove extra column
        disb_df.drop(
            columns=["Loan Type_ytd"],
            inplace=True
        )

        # =====================================================
        # SHOW DATA
        # =====================================================
        tab1, tab2 = st.tabs(
            ["📊 Mapped Data", "⚠ Unmatched"]
        )

        with tab1:

            st.write(
                f"Total disbursed: {len(disb_df)}"
            )

            unique_disb = (
                disb_df
                .groupby("Loan Type")
                .size()
                .reset_index(name="Count")
                .sort_values(by="Loan Type", ascending=True)
            )

            st.dataframe(
                unique_disb,
                use_container_width=False
            )

            disb_df = disb_df.sort_values(
                by="Loan Type",
                ascending=True
            )

            st.dataframe(
                disb_df,
                use_container_width=True
            )

        # =====================================================
        # UNMATCHED
        # =====================================================
        unmatched = disb_df[
            disb_df["Loan Type"].isna()
        ]

        with tab2:

            st.write(
                f"Unmatched rows: {len(unmatched)}"
            )

            st.dataframe(
                unmatched,
                use_container_width=True
            )

        # =====================================================
        # CONVERT AMOUNT TO NUMERIC
        # =====================================================
        if "Amount" in disb_df.columns:

            disb_df["Amount"] = pd.to_numeric(
                disb_df["Amount"],
                errors="coerce"
            )

        # =====================================================
        # DOWNLOAD EXCEL
        # =====================================================
        buffer = io.BytesIO()

        with pd.ExcelWriter(
            buffer,
            engine="openpyxl"
        ) as writer:

            disb_df.to_excel(
                writer,
                index=False,
                sheet_name="Mapped_Data"
            )

        buffer.seek(0)

        st.download_button(
            label="📥 Download Result",
            data=buffer,
            file_name="updated_disbursement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ Mapping Completed!")