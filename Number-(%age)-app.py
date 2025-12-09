import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Not Met Value Sorter", layout="centered")
st.title("🧮 Not Met Value Sorter")

uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xlsx"]
)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    # Sheet selection
    sheet_name = st.selectbox(
        "Select Sheet",
        xls.sheet_names
    )

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    st.subheader("Preview Data")
    st.dataframe(df.head())

    # Column selections
    col_name_column = st.selectbox(
        "Select column for Cancer / Type name",
        df.columns
    )

    outside_bracket_col = st.selectbox(
        "Select column for value OUTSIDE brackets (Not Met Count)",
        df.columns
    )

    inside_bracket_col = st.selectbox(
        "Select column for value INSIDE brackets (%)",
        df.columns
    )

    # ✅ Decimal selection
    decimals = st.selectbox(
        "Select number of decimal places for percentage",
        options=[0, 1, 2, 3, 4]
    )

    if st.button("Generate Output Excel"):

        def format_percentage(val):
            try:
                return f"{float(val):.{decimals}f}"
            except:
                return ""

        output_df = pd.DataFrame()
        output_df["Type of Cancer"] = df[col_name_column]

        output_df["Number of Not Met Cases (Percentage)"] = (
            df[outside_bracket_col].astype(str)
            + " ("
            + df[inside_bracket_col].apply(format_percentage)
            + "%)"
        )

        # Write to Excel in memory
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output_df.to_excel(
                writer,
                index=False,
                sheet_name="Not Met Summary"
            )

        st.success("✅ Excel generated successfully")

        st.download_button(
            label="⬇ Download Excel",
            data=buffer.getvalue(),
            file_name="Not_Met_Value_Sorter_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
