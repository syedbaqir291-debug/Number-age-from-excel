import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Formatter: Not Met (Percentage) Generator")

# Fixed category order and possible variants
category_mapping = {
    "Haematological": ["Haematological", "Haematological malignancies"],
    "Gynecological": ["Gynecological", "Gynecological Tumors", "Gynecological malignancies"],
    "Urological": ["Urological", "Urological Tumors"],
    "Neurological": ["Neurological", "Neurological malignancies"],
    "Breast": ["Breast", "Breast cancer"],
    "Pulmonary": ["Pulmonary", "Pulmonary malignancies"],
    "Gastrointestinal": ["Gastrointestinal", "Gastrointestinal malignancies"],
    "Head & Neck": ["Head & Neck", "Head and Neck", "Head-Neck", "Head & neck", "Head-neck", "Head and neck"],
    "Thyroid": ["Thyroid", "Thyroid cancers"],
    "Sarcoma": ["Sarcoma"],
    "Retinoblastoma": ["Retinoblastoma"],
    "Other rare tumors": ["Other rare tumors", "Others", "Non-specific", "non-specific", "Non Specific"]
}

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    # Load Excel file
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select sheet", xl.sheet_names)
    df = xl.parse(sheet_name)

    st.write("Preview of sheet:")
    st.dataframe(df.head())

    # Step 2: Select columns
    category_col = st.selectbox("Select column for categories", df.columns)
    outside_col = st.selectbox("Select column for outside-bracket values", df.columns)
    inside_col = st.selectbox("Select column for inside-bracket values", df.columns)

    # Step 3: Decimal places
    decimal_place = st.selectbox("Select decimal places for inside-bracket value", [0, 1, 2, 3])

    if st.button("Generate Excel"):
        # Format the new column
        df["Not Met (Non-compliance %)"] = df.apply(
            lambda row: f"{row[outside_col]} ({round(row[inside_col], decimal_place)}%)", axis=1
        )

        # Create result dataframe following fixed order & variants
        result_list = []
        for fixed_cat, variants in category_mapping.items():
            # Search for any variant in the dataframe
            match = df[df[category_col].str.contains('|'.join(variants), case=False, na=False)]
            if not match.empty:
                result_list.append({
                    category_col: fixed_cat,
                    "Not Met (Non-compliance %)": match.iloc[0]["Not Met (Non-compliance %)"]
                })
            else:
                result_list.append({
                    category_col: fixed_cat,
                    "Not Met (Non-compliance %)": "-"
                })

        result_df = pd.DataFrame(result_list)

        # Save to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Formatted Output")
        processed_data = output.getvalue()

        # Download button
        st.download_button(
            label="Download formatted Excel",
            data=processed_data,
            file_name="Formatted_NotMet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Excel generated successfully!")
