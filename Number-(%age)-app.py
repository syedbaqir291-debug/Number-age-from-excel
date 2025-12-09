import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Not Met Value Sorter", layout="centered")
st.title("🧮 Not Met Value Sorter")

uploaded_file = st.file_uploader("Upload Excel Workbook", type=["xlsx"])

# ✅ Required output order
FINAL_ORDER = [
    "Haematological",
    "Gynecological",
    "Urological",
    "Neurological",
    "Breast",
    "Pulmonary",
    "Gastrointestinal",
    "Head & Neck",
    "Thyroid",
    "Sarcoma",
    "Retinoblastoma",
    "Other rare tumors"
]

def normalize_cancer(name):
    name = str(name).lower()

    if "haemat" in name:
        return "Haematological"
    if "gyn" in name:
        return "Gynecological"
    if "uro" in name:
        return "Urological"
    if "neuro" in name:
        return "Neurological"
    if "breast" in name:
        return "Breast"
    if "pulmo" in name or "lung" in name:
        return "Pulmonary"
    if "gastro" in name or "gi" in name:
        return "Gastrointestinal"
    if "head" in name or "neck" in name:
        return "Head & Neck"
    if "thyroid" in name:
        return "Thyroid"
    if "sarcoma" in name:
        return "Sarcoma"
    if "retino" in name:
        return "Retinoblastoma"
    if "non" in name or "rare" in name or "specific" in name:
        return "Other rare tumors"

    # fallback
    return "Other rare tumors"

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    st.subheader("Preview Data")
    st.dataframe(df.head())

    cancer_col = st.selectbox("Select cancer/type column", df.columns)
    outside_col = st.selectbox("Value OUTSIDE brackets (Not Met)", df.columns)
    inside_col = st.selectbox("Value INSIDE brackets (%)", df.columns)

    decimals = st.selectbox(
        "Decimal places for percentage",
        [0, 1, 2, 3, 4]
    )

    if st.button("Generate Output Excel"):

        df["Type of Cancer"] = df[cancer_col].apply(normalize_cancer)

        df["Formatted Value"] = (
            df[outside_col].astype(str)
            + " ("
            + df[inside_col].astype(float).round(decimals).astype(str)
            + "%)"
        )

        # ✅ Aggregate in case multiple rows map to same cancer
        final_df = (
            df.groupby("Type of Cancer", as_index=False)
              .agg({"Formatted Value": "first"})
        )

        # ✅ Apply forced sorting
        final_df["Type of Cancer"] = pd.Categorical(
            final_df["Type of Cancer"],
            categories=FINAL_ORDER,
            ordered=True
        )
        final_df = final_df.sort_values("Type of Cancer")

        final_df.rename(
            columns={"Formatted Value": "Number of Not Met Cases (Percentage)"},
            inplace=True
        )

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final_df.to_excel(
                writer,
                index=False,
                sheet_name="Not Met Summary"
            )

        st.success("✅ Excel generated with fixed order & fuzzy matching")

        st.download_button(
            "⬇ Download Excel",
            buffer.getvalue(),
            file_name="Not_Met_Value_Sorter_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
