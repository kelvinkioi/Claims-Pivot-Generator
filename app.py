import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Pivot Table Generator", layout="centered")
st.title("📊 Pivot Table Generator")

# ------------------------------
# Step 1: Upload and Cache the File
# ------------------------------
uploaded_file = st.file_uploader(
    "Upload your Excel file (must contain columns: SCHEME, TRANSACTION DATE, BENEFIT, AMOUNT, COUNT, UNIQUE COUNT)",
    type=["xlsx"]
)

if uploaded_file:
    # Use session_state to avoid re-reading file on every interaction
    if "df" not in st.session_state:
        with st.spinner("Reading file..."):
            try:
                st.session_state.df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
                st.success("✅ File loaded successfully!")
            except Exception as e:
                st.error(f"Failed to read the Excel file: {e}")
                st.stop()
    
    df = st.session_state.df

    # Check if required columns exist
    required_cols = ['SCHEME', 'TRANSACTION DATE', 'BENEFIT', 'AMOUNT', 'COUNT', 'UNIQUE COUNT']
    if not all(col in df.columns for col in required_cols):
        st.error(f"Missing required columns. Expected columns: {required_cols}")
        st.stop()
    
    # Convert TRANSACTION DATE to datetime format only once
    df['TRANSACTION DATE'] = pd.to_datetime(df['TRANSACTION DATE'], errors='coerce')
    
    # ------------------------------
    # Step 2: Collect User Inputs in a Form
    # ------------------------------
    # Get all unique schemes from the file.
    unique_schemes = sorted(df['SCHEME'].dropna().unique().tolist())
    
    with st.form("input_form"):
        st.info("Select the schemes and set the date range for each (or choose to ignore dates).")
        
        # Multiselect for schemes
        selected_schemes = st.multiselect(
            "Search and select schemes:",
            options=unique_schemes,
            help="Select one or more schemes to generate pivots for."
        )
        
        # Dictionary to hold per-scheme date filter settings
        scheme_date_filters = {}
        if selected_schemes:
            st.markdown("### Configure Date Filters for Each Selected Scheme")
            # For each selected scheme, create an expander with date inputs and an option to ignore dates.
            for scheme in selected_schemes:
                with st.expander(f"Date settings for: **{scheme}**", expanded=True):
                    ignore_dates = st.checkbox(f"Process {scheme} without date filter", key=f"{scheme}_ignore")
                    if not ignore_dates:
                        # Date inputs for the scheme.
                        start_date = st.date_input(
                            f"Start Date for {scheme}",
                            key=f"{scheme}_start",
                            value=datetime.today()
                        )
                        end_date = st.date_input(
                            f"End Date for {scheme}",
                            key=f"{scheme}_end",
                            value=datetime.today()
                        )
                        scheme_date_filters[scheme] = {
                            "apply_dates": True,
                            "start_date": pd.to_datetime(start_date),
                            "end_date": pd.to_datetime(end_date)
                        }
                    else:
                        scheme_date_filters[scheme] = {
                            "apply_dates": False,
                            "start_date": None,
                            "end_date": None
                        }
        else:
            st.warning("Please select at least one scheme.")
        
        # Submit button within the form triggers heavy processing
        submitted = st.form_submit_button("Generate Pivot Tables")
    
    # ------------------------------
    # Step 3: Process and Generate Pivots (Only on Submit)
    # ------------------------------
    if submitted:
        # Validate date ranges
        for scheme, settings in scheme_date_filters.items():
            if settings["apply_dates"] and (settings["start_date"] > settings["end_date"]):
                st.error(f"For scheme {scheme}: Start Date must be before End Date.")
                st.stop()
        
        with st.spinner("Generating pivot tables. This might take a moment..."):
            output = BytesIO()
            workbook = openpyxl.Workbook()
            # Remove the default sheet created by openpyxl.
            default_sheet = workbook.active
            workbook.remove(default_sheet)
            
            # Process each selected scheme.
            for scheme in selected_schemes:
                scheme_df = df[df['SCHEME'] == scheme].copy()
                settings = scheme_date_filters.get(scheme, {"apply_dates": False})
                if settings["apply_dates"]:
                    scheme_df = scheme_df[
                        (scheme_df['TRANSACTION DATE'] >= settings["start_date"]) &
                        (scheme_df['TRANSACTION DATE'] <= settings["end_date"])
                    ]
                if scheme_df.empty:
                    st.warning(f"No data found for scheme '{scheme}' with the applied date filter. Skipping.")
                    continue
                
                scheme_df = scheme_df.sort_values(by='TRANSACTION DATE')
                scheme_df['TRANSACTION DATE NORMALIZED'] = scheme_df['TRANSACTION DATE'].dt.strftime('%m/%Y')
                
                # Create Pivot Table 1: Benefit by Amount
                pivot1 = pd.pivot_table(
                    scheme_df,
                    values='AMOUNT',
                    index='TRANSACTION DATE NORMALIZED',
                    columns='BENEFIT',
                    aggfunc='sum',
                    margins=True,
                    margins_name='Grand Total'
                )
                # Create Pivot Table 2: Benefit by Count
                pivot2 = pd.pivot_table(
                    scheme_df,
                    values='COUNT',
                    index='TRANSACTION DATE NORMALIZED',
                    columns='BENEFIT',
                    aggfunc='sum',
                    margins=True,
                    margins_name='Grand Total'
                )
                # Create Pivot Table 3: Unique Count
                pivot3 = pd.pivot_table(
                    scheme_df,
                    values='UNIQUE COUNT',
                    index='TRANSACTION DATE NORMALIZED',
                    aggfunc='sum',
                    margins=True,
                    margins_name='Grand Total'
                )
                
                # Create a new sheet for the scheme (limit to 31 characters)
                sheet_name = scheme[:31]
                if sheet_name in workbook.sheetnames:
                    counter = 1
                    new_sheet_name = f"{sheet_name} {counter}"
                    while new_sheet_name in workbook.sheetnames:
                        counter += 1
                        new_sheet_name = f"{sheet_name} {counter}"
                    sheet_name = new_sheet_name
                sheet = workbook.create_sheet(sheet_name)
                
                # Dynamically generate headers from unique BENEFIT values
                dynamic_benefits = sorted(scheme_df['BENEFIT'].unique())
                headers = ['TRANSACTION DATE NORMALIZED'] + dynamic_benefits + ['Grand Total']
                
                # Write Pivot Table 1 (Benefit by Amount)
                sheet.cell(row=1, column=1, value="Benefit by Amount")
                for col_num, header in enumerate(headers, start=1):
                    sheet.cell(row=2, column=col_num, value=header)
                for r_idx, (index, row) in enumerate(pivot1.iterrows(), start=3):
                    sheet.cell(row=r_idx, column=1, value=index)
                    for c_idx, col in enumerate(headers[1:], start=2):
                        sheet.cell(row=r_idx, column=c_idx, value=row.get(col, 0))
                
                # Write Pivot Table 2 (Benefit by Count)
                start_row_second = pivot1.shape[0] + 5
                sheet.cell(row=start_row_second, column=1, value="Benefit by Count")
                for col_num, header in enumerate(headers, start=1):
                    sheet.cell(row=start_row_second + 1, column=col_num, value=header)
                for r_idx, (index, row) in enumerate(pivot2.iterrows(), start=start_row_second + 2):
                    sheet.cell(row=r_idx, column=1, value=index)
                    for c_idx, col in enumerate(headers[1:], start=2):
                        sheet.cell(row=r_idx, column=c_idx, value=row.get(col, 0))
                
                # Write Pivot Table 3 (Unique Count)
                start_row_third = start_row_second + pivot2.shape[0] + 5
                sheet.cell(row=start_row_third, column=1, value="Number of Lives (Unique Count)")
                sheet.cell(row=start_row_third + 1, column=1, value="TRANSACTION DATE NORMALIZED")
                sheet.cell(row=start_row_third + 1, column=2, value="UNIQUE COUNT")
                for r_idx, row in enumerate(pivot3.iterrows(), start=start_row_third + 2):
                    sheet.cell(row=r_idx, column=1, value=row[0])
                    sheet.cell(row=r_idx, column=2, value=row[1]['UNIQUE COUNT'])
            
            workbook.save(output)
            output.seek(0)
        
        st.success("✅ Pivot tables generated successfully!")
        st.download_button(
            label="📥 Download Pivot Excel File",
            data=output.getvalue(),
            file_name="Pivot_Tables_By_Scheme.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    # Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center;'>Developed with ❤️ by "
    "<a href='https://github.com/kelvinkioi/' target='_blank'>Kelvin Kioi</a></div>",
    unsafe_allow_html=True)
