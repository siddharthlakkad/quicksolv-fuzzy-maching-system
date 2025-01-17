import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

# Set page configuration
st.set_page_config(page_title="Optimized Fuzzy Matching Tool", page_icon=":mag:", layout="centered")

# Add a company logo
st.image("image/logo1.png", width=500)

# Custom CSS for animations and styling
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

# Title with animation
st.markdown(
    """
    <div style="text-align: center;">
        <h1 style="color: #3498db;">Fuzzy Matching Tool</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# File uploaders for source and destination files
file1 = st.file_uploader("Select Source File", type=["xlsx"], key="file1_uploader")
file2 = st.file_uploader("Select Destination File", type=["xlsx"], key="file2_uploader")

if file1 and file2:
    # Load and display file details
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    col1_name = st.selectbox("Select the column from the Source file to match:", df1.columns, key="col1_selectbox")
    col2_name = st.selectbox("Select the column from the Destination file to match:", df2.columns, key="col2_selectbox")

    # Set fuzzy match threshold
    threshold = st.slider("Set the matching threshold", min_value=0, max_value=100, value=80, key="threshold_slider")
    model_name = st.selectbox("Select Matching Model", ["Ratio", "Partial Ratio", "Token Sort Ratio"], key="model_selectbox")

    # Multiselect for columns to display in the output
    source_columns_to_display = st.multiselect("Select columns from Source file to display in the result:", df1.columns, default=df1.columns)
    destination_columns_to_display = st.multiselect("Select columns from Destination file to display in the result:", df2.columns, default=df2.columns)

    def optimized_fuzzy_match(source_df, target_df, source_col, target_col, threshold, model):
        match_func = {
            "Ratio": fuzz.ratio,
            "Partial Ratio": fuzz.partial_ratio,
            "Token Sort Ratio": fuzz.token_sort_ratio
        }[model]

        matches = []
        # Filter out blank or NaN values to speed up processing
        source_values = source_df[source_col].dropna().astype(str)
        target_values = target_df[target_col].dropna().astype(str)

        # Use process.extract to get the best matches quickly
        for idx, src_value in source_values.items():
            results = process.extract(
                src_value, target_values, scorer=match_func, limit=5  # Adjust limit for balance between performance and match accuracy
            )
            for match_value, score, match_idx in results:
                if score >= threshold:
                    # Compile source and target data with scores for each match
                    source_row = source_df.loc[idx, source_columns_to_display].to_dict()
                    dest_row = target_df.loc[match_idx, destination_columns_to_display].to_dict()
                    match = {
                        **{f"Source_{key}": value for key, value in source_row.items()},
                        **{f"Destination_{key}": value for key, value in dest_row.items()},
                        "Similarity": score
                    }
                    matches.append(match)

        return pd.DataFrame(matches)

    if st.button("Perform Optimized Fuzzy Matching", key="optimized_match_button"):
        matches_df = optimized_fuzzy_match(df1, df2, col1_name, col2_name, threshold, model_name)

        # Display results
        st.markdown(
            """
            <div class="result-container">
                <h2 class="result-title">Optimized Fuzzy Matching Results</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

        if not matches_df.empty:
            st.dataframe(matches_df)
        else:
            st.write("No matches found.")
