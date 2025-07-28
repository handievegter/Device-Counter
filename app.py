import streamlit as st
import pandas as pd

st.title("📊 Excel Processor App")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)
    
    st.subheader("Preview of uploaded file:")
    st.dataframe(df)

    # Dummy processing: add a new column
    df["Processed"] = "✅"

    # Download processed file
    @st.cache_data
    def convert_df_to_excel(df):
        return df.to_excel(index=False, engine='openpyxl')

    st.download_button(
        label="📥 Download Processed Excel",
        data=convert_df_to_excel(df),
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
