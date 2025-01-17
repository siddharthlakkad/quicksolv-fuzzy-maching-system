
import pandas as pd
import streamlit as st
from fuzzywuzzy import fuzz
import Levenshtein as lev

# Set page configuration
st.set_page_config(page_title="Fuzzy Matching Tool", page_icon=":mag:", layout="centered")

# Add a company logo
st.image("image/logo1.png", width=500)

# Custom CSS for animations, touch, and click effects, including styling for the custom HTML UI
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
    .model-info {
        font-size: 16px;
        color: #000;
        font-weight: bold;
        background-color: #d1ecf1;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
    }
    .model-description {
        margin-top: 5px;
        font-style: italic;
        font-weight: normal;
        background-color: #ffefba;
        padding: 10px;
        border-radius: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title with animation
st.markdown(
    """
    <div style="text-align: center;">
        <h1 style="color: #3498db;">Fuzzy Matching Excel Columns</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# Dictionary containing model efficiency levels and descriptions
model_info = {
    "Ratio": {
        "efficiency": "Moderate Efficiency",
        "description": "General purpose string comparison that checks for character similarity."
    },
    "Partial Ratio": {
        "efficiency": "High Efficiency",
        "description": "Works well when substrings match, even if full strings are not exactly the same."
    },
    "Token Sort Ratio": {
        "efficiency": "Moderate Efficiency",
        "description": "Effective when comparing strings where words are rearranged."
    },
    "Levenshtein": {
        "efficiency": "High Efficiency",
        "description": "Measures the number of single-character edits required to change one string to another."
    }
}

# File uploader for the first Excel file
file1 = st.file_uploader("Select Source File", type=["xlsx"], key="file1_uploader")

if file1 is not None:
    df1 = pd.read_excel(file1)
    st.write("Columns in the Source file:", list(df1.columns))
    col1_name = st.selectbox("Select the column from the first file to match:", df1.columns, key="col1_selectbox")

# File uploader for the second Excel file
file2 = st.file_uploader("Select Destination File", type=["xlsx"], key="file2_uploader")

if file2 is not None:
    df2 = pd.read_excel(file2)
    st.write("Columns in the Destination file:", list(df2.columns))
    col2_name = st.selectbox("Select the column from the Destination file to match:", df2.columns, key="col2_selectbox")

if file1 is not None and file2 is not None and col1_name and col2_name:
    col1 = df1[col1_name]
    col2 = df2[col2_name]

    # Set fuzzy match threshold
    threshold = st.slider("Set the matching threshold", min_value=0, max_value=100, value=80, key="threshold_slider")

    # Function to select the matching model
    def select_matching_model(model_name):
        if model_name == "Levenshtein":
            return lambda x, y: 100 - lev.distance(x, y) * 100 / max(len(x), len(y))
        elif model_name == "Partial Ratio":
            return fuzz.partial_ratio
        elif model_name == "Token Sort Ratio":
            return fuzz.token_sort_ratio
        else:  # Default is fuzz.ratio
            return fuzz.ratio

    # Model selection dropdown
    model_name = st.selectbox(
        "Select Matching Model",
        ["Ratio", "Partial Ratio", "Token Sort Ratio", "Levenshtein"],
        key="model_selectbox"
    )

    # Display model efficiency level and description with bold and highlighted text
    st.markdown(
        f"""
        <div class="model-info">
            <strong>Model Efficiency: </strong> {model_info[model_name]["efficiency"]}
        </div>
        <div class="model-description">
            <strong>Description: </strong> {model_info[model_name]["description"]}
        </div>
        """,
        unsafe_allow_html=True
    )

    # Multiselect for columns to display in the output
    source_columns_to_display = st.multiselect(
        "Select columns from Source file to display in the result:",
        options=df1.columns,
        default=df1.columns
    )

    destination_columns_to_display = st.multiselect(
        "Select columns from Destination file to display in the result:",
        options=df2.columns,
        default=df2.columns
    )

    # Define the fuzzy matching function with selected model and display columns
    def fuzzy_match(df1, df2, col1_name, col2_name, threshold, model_name, source_cols, dest_cols):
        match_func = select_matching_model(model_name)
        matches = []

        for i, val1 in df1[col1_name].items():
            for j, val2 in df2[col2_name].items():
                ratio = match_func(str(val1), str(val2))  # Convert values to strings if they are not
                if ratio >= threshold:
                    # Filter the source and destination rows based on selected columns
                    source_row = df1.loc[i, source_cols].to_dict()
                    dest_row = df2.loc[j, dest_cols].to_dict()
                    match = {
                        **{f"Source_{key}": value for key, value in source_row.items()},
                        **{f"Destination_{key}": value for key, value in dest_row.items()},
                        'Similarity': ratio
                    }
                    matches.append(match)
        return pd.DataFrame(matches)

    if st.button("Perform Fuzzy Matching", key="match_button"):
        matches_df = fuzzy_match(df1, df2, col1_name, col2_name, threshold, model_name, source_columns_to_display, destination_columns_to_display)
        
        # Custom section to display results
        st.markdown(
            """
            <div class="result-container">
                <h2 class="result-title">Fuzzy Matching Results</h2>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Display the results inside the custom container
        st.dataframe(matches_df)
