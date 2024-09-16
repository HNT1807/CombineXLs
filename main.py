import streamlit as st
import pandas as pd
import io
import uuid
import base64

st.set_page_config(page_title="COMBINE XLs AND KEEP FULLS", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .main-title { font-size: 32px; font-weight: bold; text-align: center; margin-bottom: 30px; }
    .download-button { text-align: center; margin-top: 20px; }
    .stApp > header { display: none !important; }
    .block-container { max-width: 1000px; padding-top: 1rem; padding-bottom: 10rem; }
    .custom-button {
        background-color: #4CAF50;
        border: none;
        color: white;
        padding: 15px 32px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
        transition: all 0.3s ease 0s;
        box-shadow: 0 8px 15px rgba(0, 0, 0, 0.1);
    }
    .custom-button:hover {
        background-color: #45a049;
        box-shadow: 0 15px 20px rgba(0, 0, 0, 0.2);
        transform: translateY(-7px);
    }
</style>
""", unsafe_allow_html=True)


def check_empty_values(df, filename):
    non_empty_rows = df.apply(lambda row: not row.isna().all(), axis=1)
    empty_t_values = df.loc[non_empty_rows, df.columns[19]].isna()
    if empty_t_values.any():
        st.error(
            f"Error: File '{filename}' has {empty_t_values.sum()} empty value(s) in column T for non-empty rows. Processing stopped.")
        return False
    return True


def process_excel_files(uploaded_files):
    combined_df = None
    for i, file in enumerate(uploaded_files):
        df = pd.read_excel(file)

        if not check_empty_values(df, file.name):
            return None

        if i == 0:
            combined_df = df
        else:
            combined_df = pd.concat([combined_df, df.iloc[1:]], ignore_index=True)

    combined_df = combined_df[combined_df.iloc[:, 19].astype(str).str.contains("Full", case=False, na=False)]

    return combined_df


def get_binary_file_downloader_html(df):
    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="combined_excel_{uuid.uuid4()}.xlsx"><button class="custom-button">Download Combined Excel File</button></a>'


# Main app layout
st.markdown("<h1 class='main-title'>COMBINE XLs AND KEEP FULLS</h1>", unsafe_allow_html=True)

# Create a centered column
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

    if uploaded_files:
        if st.button("Combine and Process Files"):
            with st.spinner("Processing files..."):
                combined_df = process_excel_files(uploaded_files)
                if combined_df is not None:
                    st.success(f"Files combined successfully! Total rows after processing: {len(combined_df)}")

                    st.markdown("<div class='download-button'>", unsafe_allow_html=True)
                    st.markdown(get_binary_file_downloader_html(combined_df), unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                else:
                    st.error("Processing stopped due to empty values in column T. Please fix the issues and try again.")
    else:
        st.info("Please upload Excel (.xlsx) files to combine.")
