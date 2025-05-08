import streamlit as st

import pandas as pd

from rapidfuzz import process, fuzz

import time

 

# Function to preprocess text (remove extra spaces and convert to lowercase)

def clean_text(text):

    return str(text).strip().lower() if pd.notna(text) else ""

 

# Function to find the best match for each Range 2 customer

def match_customers_optimized(range1, range2):

    start_time = time.time()

 

    # Convert name & address columns to lowercase

    range1["clean_name"] = range1["cus_name"].apply(clean_text)

    range1["clean_address"] = range1["cus_address"].apply(clean_text)

    range2["clean_name"] = range2["cus_name"].apply(clean_text)

    range2["clean_address"] = range2["cus_address"].apply(clean_text)

 

    matches = []

 

    # Create a combined key for Range 1 to speed up searches

    range1_dict = range1.set_index("clean_name")["clean_address"].to_dict()

 

    for idx2, row2 in range2.iterrows():

        query = row2["clean_name"]

       

        # Find the best match using process.extractOne

        best_match, best_score, best_index = process.extractOne(

            query, range1_dict.keys(), scorer=fuzz.ratio

        )

 

        # If a match is found, compare addresses as well

        if best_match and best_score > 60:

            address_score = fuzz.ratio(range1_dict[best_match], row2["clean_address"])

            overall_score = (best_score + address_score) / 2

 

            if overall_score > 60:  # Only consider significant matches

                matches.append((best_match, range1_dict[best_match], row2["cus_name"], row2["cus_address"], overall_score))

 

    print(f"Matching completed in {time.time() - start_time:.2f} seconds.")

    return matches

 

# Streamlit UI

st.title("Fast Customer Matching Tool")

 

st.write("Upload two Excel files with customer data to find similar customers.")

 

file1 = st.file_uploader("Upload Range 1 Excel File", type=["xlsx"])

file2 = st.file_uploader("Upload Range 2 Excel File", type=["xlsx"])

 

if file1 and file2:

    range1 = pd.read_excel(file1)

    range2 = pd.read_excel(file2)

 

    required_columns = {"cus_name", "cus_address"}

    if not required_columns.issubset(range1.columns) or not required_columns.issubset(range2.columns):

        st.error("Both files must contain 'Name' and 'Address' columns.")

    else:

        with st.spinner("Matching customers, please wait..."):

            matches = match_customers_optimized(range1, range2)

 

        matched_df = pd.DataFrame(matches, columns=["Range1_Name", "Range1_Address", "Range2_Name", "Range2_Address", "Similarity_Score"])

 

        st.write("### Matching Results")

        st.dataframe(matched_df)

 

        if not matched_df.empty:

            st.download_button(

                label="Download Results as Excel",

                data=matched_df.to_excel(index=False, engine="openpyxl"),

                file_name="matched_customers.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            )