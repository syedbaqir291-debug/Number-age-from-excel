import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Not Met Value Sorter", layout="centered")
st.title("🧮 Not Met Value Sorter")

# Final forced order
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
    if "overall" in name or "skm" in name:
        return None
    if "haemat" in name or "hemat" in name:
        return "Haematological"
    if "gyne" in name or "gyn" in name:
        return "Gynecological"
    if "uro" in name:
        return "Urological"
    if "neuro" in name:
        return "Neurological"
    if "breast" in name:
        return "Breast"
    if "pulmo" in name or "lung" in name:
        return "Pulmonary"
    if "gastro" in name:
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
    return "Other rare tumors"

uploaded_file = st.file_uploader("Upload Excel Workbook", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    st.subheader("Preview")
    st.dataframe(df.head())

    cancer_col = st.selectbox("Select Cancer / Type column", df.columns)
    outside_col = st.selectbox("Value OUTSIDE brackets (Not Met Count)", df.columns)
    inside_col = st.selectbox("Value INSIDE brackets (%)", df.columns)
    decimals = st.selectbox("Decimal places for percentage", [0, 1, 2, 3, 4])

    if st.button("Generate Output Excel"):
        # Normalize cancer type
        df["Type of Cancer"] = df[cancer_col].apply(normalize_cancer)
        df = df[df["Type of Cancer"].notna()]

        # Ensure numeric values
        df[outside_col] = pd.to_numeric(df[outside_col], errors='coerce').fillna(0)
        df[inside_col] = pd.to_numeric(df[inside_col], errors='coerce').fillna(0)

        # Aggregate sum for Not Met, mean for % inside brackets
        agg_df = df.groupby("Type of Cancer", as_index=False).agg({
            outside_col: "sum",
            inside_col: "mean"
        })

        # Format as "Count (Percentage)"
        agg_df["Number of Not Met Cases (Percentage)"] = (
            agg_df[outside_col].astype(int).astype(str)
            + " ("
            + agg_df[inside_col].round(decimals).astype(str)
            + "%)"
        )

        # Merge with FINAL_ORDER to include all categories
        final_df = pd.DataFrame({"Type of Cancer": FINAL_ORDER})
        final_df = final_df.merge(
            agg_df[["Type of Cancer", "Number of Not Met Cases (Percentage)"]],
            on="Type of Cancer", how="left"
        )
        final_df["Number of Not Met Cases (Percentage)"].fillna("0 (0%)", inplace=True)

        # Export to Excel
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Not Met Summary")

        st.success("✅ Excel generated successfully")
        st.download_button(
            "⬇ Download Excel",
            buffer.getvalue(),
            file_name="Not_Met_Value_Sorter_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
