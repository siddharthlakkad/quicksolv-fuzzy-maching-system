import time
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

# Set page configuration
st.set_page_config(page_title="Optimized Fuzzy Matching Tool", page_icon=":mag:", layout="centered")

# Add a company logo
st.image("image/logo1.png", width=600)

# Custom CSS for styling
st.markdown(
    """
    <style>
    .result-container {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #3498db;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        margin-top: 20px;
    }
    .result-title {
        color: #2ecc71;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title
st.markdown("<h1 style='text-align: center; color: #3498db;'>Fuzzy Matching Tool</h1>", unsafe_allow_html=True)

# File uploaders
file1 = st.file_uploader("Select Source File", type=["xlsx"], key="file1_uploader")
file2 = st.file_uploader("Select Destination File", type=["xlsx"], key="file2_uploader")

if file1 and file2:
    source_sheets = pd.ExcelFile(file1).sheet_names
    destination_sheets = pd.ExcelFile(file2).sheet_names
    
    source_sheet = st.selectbox("Select Sheet from Source File", source_sheets, key="source_sheet_selectbox")
    destination_sheet = st.selectbox("Select Sheet from Destination File", destination_sheets, key="destination_sheet_selectbox")
    
    df1 = pd.read_excel(file1, sheet_name=source_sheet)
    df2 = pd.read_excel(file2, sheet_name=destination_sheet)
    
    col1_name = st.selectbox("Select column from Source file:", df1.columns, key="col1_selectbox")
    col2_name = st.selectbox("Select column from Destination file:", df2.columns, key="col2_selectbox")
    
    threshold = st.slider("Set matching threshold", min_value=0, max_value=100, value=80, key="threshold_slider")
    model_name = st.selectbox("Select Matching Model", ["Ratio", "Partial Ratio", "Token Sort Ratio"], key="model_selectbox")

    model_descriptions = {
        "Ratio": "Compares full strings using Levenshtein distance. Best for exact or near-exact matches.",
        "Partial Ratio": "Compares substrings within longer texts. Useful when extra words exist.",
        "Token Sort Ratio": "Sorts words before comparison. Best when word order varies."
    }
    
    st.markdown(f"**Model Description:** {model_descriptions[model_name]}")

    source_columns_to_display = st.multiselect("Select columns from Source file:", df1.columns, default=df1.columns)
    destination_columns_to_display = st.multiselect("Select columns from Destination file:", df2.columns, default=df2.columns)

    def optimized_fuzzy_match(source_df, target_df, source_col, target_col, threshold, model):
        start_time = time.time()  
        match_func = {"Ratio": fuzz.ratio, "Partial Ratio": fuzz.partial_ratio, "Token Sort Ratio": fuzz.token_sort_ratio}[model]

        matches = []
        source_values = source_df[source_col].dropna().astype(str)
        target_values = target_df[target_col].dropna().astype(str)

        total = len(source_values)  # Total items to process
        progress_bar = st.progress(0)  # Initialize progress bar
        status_text = st.empty()  # Empty status text for updates

        for i, (idx, src_value) in enumerate(source_values.items()):
            results = process.extract(src_value, target_values, scorer=match_func, limit=5)

            for match_value, score, match_idx in results:
                if score >= threshold:
                    source_row = source_df.loc[idx, source_columns_to_display].to_dict()
                    dest_row = target_df.loc[match_idx, destination_columns_to_display].to_dict()
                    match = {**{f"Source_{key}": value for key, value in source_row.items()},
                             **{f"Destination_{key}": value for key, value in dest_row.items()},
                             "Similarity": score}
                    matches.append(match)

            # Update progress bar
            progress_bar.progress((i + 1) / total)

            # Estimate remaining time
            elapsed_time = time.time() - start_time
            avg_time_per_entry = elapsed_time / (i + 1)
            estimated_total_time = avg_time_per_entry * total
            remaining_time = max(0, estimated_total_time - elapsed_time)

            status_text.text(f"Processing {i + 1}/{total}... Estimated time left: {remaining_time:.2f} seconds")

        execution_time = time.time() - start_time
        progress_bar.empty()  # Remove progress bar
        status_text.text(f"✅ Matching completed in {execution_time:.2f} seconds.")
        return pd.DataFrame(matches), execution_time

    if st.button("Perform Optimized Fuzzy Matching", key="optimized_match_button"):
        matches_df, exec_time = optimized_fuzzy_match(df1, df2, col1_name, col2_name, threshold, model_name)

        st.markdown("<div class='result-container'><h2 class='result-title'>Matching Results</h2></div>", unsafe_allow_html=True)

        if not matches_df.empty:
            st.dataframe(matches_df)
            st.success(f"✅ Matching completed in {exec_time:.2f} seconds.")
        else:
            st.warning(f"⚠️ No matches found. Process completed in {exec_time:.2f} seconds.")
