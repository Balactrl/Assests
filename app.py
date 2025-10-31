import streamlit as st
import pandas as pd
import os
from io import BytesIO

# Set page config
st.set_page_config(page_title="Assets Audit System", layout="wide")

# Title of the application
st.title("Assets Audit Management System")

# Function to load data
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("assets.xlsx")
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()

def save_data(df):
    """Save the dataframe to Excel and clear the cache"""
    try:
        df.to_excel("assets.xlsx", index=False)
        # Clear the cached data to force a reload
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"Error saving data: {e}")
        return False

# Load the data
df = load_data()

# Search functionality
st.subheader("Search Assets")
search_siteid = st.text_input("Enter Site ID to search:")

if search_siteid:
    filtered_df = df[df['SITEID'].astype(str).str.contains(search_siteid, case=False)]
    if not filtered_df.empty:
        st.write("Search Results:")
        
        # Display the search results
        for index, row in filtered_df.iterrows():
            # control expander open/close via session state so we can close it after update
            exp_key = f"exp_{index}"
            if exp_key not in st.session_state:
                # open expander by default when search results appear
                st.session_state[exp_key] = True
            with st.expander(f"Site: {row['SITENAME']} (ID: {row['SITEID']})", expanded=st.session_state[exp_key]):
                col1, col2 = st.columns(2)
                
                # Display current values
                with col1:
                    st.write("Current Values:")
                    for column in df.columns:
                        st.write(f"{column}: {row[column]}")
                
                # Edit form
                with col2:
                    st.write("Edit Values:")
                    new_values = {}
                    for column in df.columns:
                        if column not in ['SITEID', 'SITENAME']:  # Don't allow editing of ID and Name
                            new_values[column] = st.text_input(
                                f"New {column}",
                                value=str(row[column]),
                                key=f"{index}_{column}"
                            )
                    
                    if st.button("Update", key=f"update_{index}"):
                        try:
                            # Update the dataframe
                            for column, value in new_values.items():
                                df.at[index, column] = value
                            
                            # Save back to Excel using the save function
                            if save_data(df):
                                st.success("Data updated successfully!")
                                # Close the expander for this item
                                try:
                                    st.session_state[exp_key] = False
                                except Exception:
                                    pass
                                # Force reload the data
                                df = load_data()
                            # Try to rerun programmatically if available; otherwise instruct user to refresh.
                            if hasattr(st, "experimental_rerun"):
                                st.experimental_rerun()
                            elif hasattr(st, "query_params"):
                                # change query params to force a rerun in newer Streamlit versions
                                # (set query params via the new property as requested)
                                try:
                                    st.query_params = {"_updated": str(pd.Timestamp.now().timestamp())}
                                except Exception:
                                    # If assigning fails, fall back to informing the user
                                    st.info("Please refresh the page to see updated data.")
                            else:
                                st.info("Please refresh the page to see updated data.")
                        except Exception as e:
                            st.error(f"Error updating data: {e}")
    else:
        st.warning("No matching records found.")

# Display full dataset
st.subheader("All Assets")
st.dataframe(df, width=1200, height=300)

# Download button for the updated Excel file
# Convert DataFrame to Excel
output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
excel_data = output.getvalue()

if st.download_button(
    label="Download Updated Excel File",
    data=excel_data,
    file_name="assets.xlsx",
    mime="application/vnd.ms-excel"
):
    st.success("File downloaded successfully!")
