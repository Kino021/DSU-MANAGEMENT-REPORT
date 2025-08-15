import streamlit as st
import pandas as pd
import numpy as np

# Streamlit page configuration
st.set_page_config(page_title="SPM Management DSU Report", layout="wide")

# Sidebar file uploader for multiple files
st.sidebar.header("Upload Excel Files")
uploaded_files = st.sidebar.file_uploader("Choose Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        # Initialize an empty list to store DataFrames
        dfs = []

        # Read each uploaded Excel file
        for uploaded_file in uploaded_files:
            df = pd.read_excel(uploaded_file)
            dfs.append(df)

        # Concatenate all DataFrames
        if dfs:
            df_combined = pd.concat(dfs, ignore_index=True)

            # Define exclude conditions
            exclude_status = ['Abort', 'LOCKED', 'UNLOCKED']
            exclude_remark_by = ['SPMADRID', 'SP MADRID']
            exclude_remark = [
                'Broadcast', 'Broken Promise', 'New files imported',
                'Updates when case reassign to another collector', 'NDF IN ICS',
                'FOR PULL OUT (END OF HANDLING PERIOD)', 'END OF HANDLING PERIOD',
                'New Assignment -', 'File Unhold'
            ]

            # Apply exclude filters with NaN handling
            df_filtered = df_combined[
                (~df_combined['Status'].str.contains('|'.join(exclude_status), case=False, na=False)) &
                (~df_combined['Remark By'].str.contains('|'.join(exclude_remark_by), case=False, na=False)) &
                (~df_combined['Remark'].str.contains('|'.join(exclude_remark), case=False, na=False))
            ]

            # Agent calculation: Unique Remark By with Call Duration > 0
            agent_df = df_filtered[df_filtered['Call Duration'].fillna(0) > 0]
            unique_agents = agent_df['Remark By'].nunique()

            # Accounts calculation: Unique Account No. for specific Remark Types
            follow_up_system = df_filtered[
                (df_filtered['Remark Type'].str.contains('Follow Up', case=False, na=False)) &
                (df_filtered['Remark By'].str.contains('System|SYSTEM', case=False, na=False))
            ]
            predictive_outgoing = df_filtered[
                df_filtered['Remark Type'].str.contains('Predictive|Outgoing', case=False, na=False)
            ]
            accounts_df = pd.concat([follow_up_system, predictive_outgoing])
            unique_accounts = accounts_df['Account No.'].nunique()

            # Dials calculation: Count rows for specific Remark Types
            dials_df = df_filtered[
                (
                    (df_filtered['Remark Type'].str.contains('Follow Up', case=False, na=False)) &
                    (df_filtered['Remark By'].str.contains('System|SYSTEM', case=False, na=False))
                ) |
                (df_filtered['Remark Type'].str.contains('Predictive|Outgoing', case=False, na=False))
            ]
            total_dials = len(dials_df)

            # Conn Unique: Unique Account No. for rows with Talk Time Duration > 0
            conn_unique_df = accounts_df[accounts_df['Talk Time Duration'].fillna(0) > 0]
            conn_unique = conn_unique_df['Account No.'].nunique()

            # Create summary table
            summary_data = {
                'Agent': [unique_agents],
                'Accounts': [unique_accounts],
                'Dials': [total_dials],
                'Conn Unique': [conn_unique]
            }
            summary_df = pd.DataFrame(summary_data)

            # Display the report
            st.header("SPM MANAGEMENT DSU REPORT")
            st.subheader("Summary Table")
            st.table(summary_df)

            # Convert DataFrames to CSV for download
            raw_csv = dials_df.to_csv(index=False)
            summary_csv = summary_df.to_csv(index=False)

            # Add download buttons
            st.subheader("Download Data")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download Raw Filtered Data (CSV)",
                    data=raw_csv,
                    file_name="spm_dsu_raw_data.csv",
                    mime="text/csv"
                )
            with col2:
                st.download_button(
                    label="Download Summary Table (CSV)",
                    data=summary_csv,
                    file_name="spm_dsu_summary.csv",
                    mime="text/csv"
                )

            # Debug option: Show filtered dials data if needed
            with st.expander("Inspect Dials Data"):
                st.write("Filtered data used for Dials calculation:")
                st.dataframe(dials_df[['Remark Type', 'Remark By', 'Account No.']])

        else:
            st.warning("No data was loaded from the uploaded files.")

    except Exception as e:
        st.error(f"An error occurred while processing the files: {str(e)}")
else:
    st.info("Please upload one or more Excel files to generate the report.")
