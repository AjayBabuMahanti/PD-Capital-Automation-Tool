import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Custom CSS for styling - Modern UI like the image
st.markdown("""
<style>
    /* Main title styling */
    .main-title {
        text-align: left;
        color: #1E3A8A;
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 10px;
        padding-bottom: 10px;
        border-bottom: 3px solid #3B82F6;
    }
    
    /* Subtitle styling */
    .subtitle {
        text-align: left;
        color: #6B7280;
        font-size: 14px;
        font-weight: 400;
        margin-bottom: 30px;
    }
    
    /* Section headers */
    .section-header {
        color: #1E3A8A;
        font-size: 18px;
        font-weight: 700;
        margin-top: 25px;
        margin-bottom: 15px;
        padding-bottom: 5px;
        border-bottom: 2px solid #E5E7EB;
    }
    
    /* Orange headers for first row */
    .orange-header {
        background-color: #FFA500 !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 11px !important;
        text-align: center !important;
        padding: 8px 4px !important;
        border: 1px solid #D1D5DB !important;
    }
    
    /* Gray headers for second row */
    .gray-header {
        background-color: #F3F4F6 !important;
        color: #111827 !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        text-align: center !important;
        padding: 8px 4px !important;
        border: 1px solid #D1D5DB !important;
    }
    
    /* Card styling for info boxes */
    .info-card {
        background-color: #F8FAFC;
        border-left: 4px solid #3B82F6;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    
    /* Success message */
    .success-box {
        background-color: #D1FAE5;
        border: 1px solid #10B981;
        border-radius: 6px;
        padding: 12px;
        color: #065F46;
        font-weight: 500;
    }
    
    /* Warning message */
    .warning-box {
        background-color: #FEF3C7;
        border: 1px solid #F59E0B;
        border-radius: 6px;
        padding: 12px;
        color: #92400E;
        font-weight: 500;
    }
    
    /* Error message */
    .error-box {
        background-color: #FEE2E2;
        border: 1px solid #EF4444;
        border-radius: 6px;
        padding: 12px;
        color: #991B1B;
        font-weight: 500;
    }
    
    /* Status badges */
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 8px;
    }
    
    .status-success {
        background-color: #D1FAE5;
        color: #065F46;
    }
    
    .status-warning {
        background-color: #FEF3C7;
        color: #92400E;
    }
    
    /* Button styling */
    .stButton > button {
        background-color: #3B82F6;
        color: white;
        font-weight: 600;
        border: none;
        padding: 12px 24px;
        border-radius: 6px;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        background-color: #2563EB;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
    
    /* Upload area styling */
    .upload-area {
        border: 2px dashed #D1D5DB;
        border-radius: 8px;
        padding: 30px;
        text-align: center;
        background-color: #F9FAFB;
        margin: 20px 0;
    }
    
    /* Data table styling */
    .dataframe {
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Sidebar styling */
    .sidebar-content {
        padding: 20px 15px;
    }
    
    /* Metric cards */
    .metric-card {
        background: white;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Highlight important text */
    .highlight {
        background-color: #E0F2FE;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
        color: #0369A1;
    }
</style>
""", unsafe_allow_html=True)

def extract_portfolio_name(filename):
    """Extract first 10 alphanumeric characters from filename as portfolio name"""
    clean_name = filename.split('/')[-1].split('\\')[-1]
    alphanumeric = ''.join(c for c in clean_name if c.isalnum())
    return alphanumeric[:10] if alphanumeric else "Portfolio"

def validate_isin(isin_code):
    """Validate ISIN code - should be 12 alphanumeric characters"""
    if pd.isna(isin_code):
        return ""
    
    isin_str = str(isin_code).strip()
    
    if len(isin_str) == 12 and isin_str.isalnum():
        if isin_str[:2].isalpha():
            return isin_str.upper()
    
    return ""

def process_single_file(input_file, portfolio_name):
    """Process a single input Excel file"""
    try:
        df_input = pd.read_excel(input_file)
    except Exception as e:
        st.error(f"Error reading {input_file.name}: {str(e)}")
        return None
    
    required_columns = [
        'ISIN Code', 'Quantity / Notional', 'Cost Price (Ref. Cur)', 
        'Reporting Current Rate Date'
    ]
    
    missing_columns = [col for col in required_columns if col not in df_input.columns]
    if missing_columns:
        st.warning(f"Missing columns in {input_file.name}: {', '.join(missing_columns)}")
        return None
    
    output_data = []
    
    for idx, row in df_input.iterrows():
        isin_code = row.get('ISIN Code', '')
        isin_value = validate_isin(isin_code)
        
        quantity = row.get('Quantity / Notional', '')
        if pd.notna(quantity):
            try:
                quantity = float(quantity)
            except:
                pass
        
        cost_price = row.get('Cost Price (Ref. Cur)', '')
        if pd.notna(cost_price):
            try:
                cost_price = float(cost_price)
            except:
                pass
        
        date_value = row.get('Reporting Current Rate Date', '')
        if pd.notna(date_value):
            try:
                if isinstance(date_value, str):
                    date_value = pd.to_datetime(date_value, errors='coerce')
                if pd.isna(date_value):
                    date_value = ""
                else:
                    date_value = date_value.strftime('%Y-%m-%d')
            except:
                date_value = str(date_value)
        else:
            date_value = ""
        
        output_row = {
            'Portfolio Name': portfolio_name,
            'Security': "",
            'Sedol': "",
            'Cusip': "",
            'ISIN': isin_value,
            'Security Name': "",
            'Position': quantity,
            'Weight': "",
            'Mkt Px': "",
            'Cost Price': cost_price,
            'As of Date': date_value,
            'New Classification': ""
        }
        
        output_data.append(output_row)
    
    if output_data:
        return pd.DataFrame(output_data)
    return None

def process_multiple_files(uploaded_files):
    """Process multiple Excel files and combine them"""
    all_outputs = []
    processed_files = 0
    failed_files = []
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        progress = (i + 1) / len(uploaded_files)
        progress_bar.progress(progress)
        status_text.text(f"Processing {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
        
        portfolio_name = extract_portfolio_name(uploaded_file.name)
        df_output = process_single_file(uploaded_file, portfolio_name)
        
        if df_output is not None:
            all_outputs.append(df_output)
            processed_files += 1
        else:
            failed_files.append(uploaded_file.name)
    
    progress_bar.empty()
    status_text.empty()
    
    if all_outputs:
        combined_df = pd.concat(all_outputs, ignore_index=True)
        return combined_df, processed_files, failed_files
    else:
        return None, processed_files, failed_files

def to_excel(df):
    """Convert DataFrame to Excel bytes with proper formatting using xlsxwriter"""
    output = BytesIO()
    
    # Use xlsxwriter for formatting
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write data starting from row 2 WITHOUT headers
        df.to_excel(writer, index=False, sheet_name='Output', startrow=2, header=False)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Output']
        
        # Create format for orange headers
        orange_format = workbook.add_format({
            'bg_color': '#FFA500',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_size': 10
        })
        
        # Create format for gray headers
        gray_format = workbook.add_format({
            'bg_color': '#D9D9D9',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_size': 10
        })
        
        # First row headers (Orange)
        first_row_headers = [
            'PORTFOLIO NAME',
            'SECURITY_ID',
            'SECURITY_ID',
            'SECURITY_ID',
            'SECURITY_ID',
            '',
            'QUANTITY',
            '"Fixed or\nDrifting Weight"',
            'Market Price',
            'Cost Price',
            'As of Date',
            'Custom grouping'
        ]
        
        # Write first row
        for col_num, value in enumerate(first_row_headers):
            worksheet.write(0, col_num, value, orange_format)
        
        # Write second row headers
        for col_num, column_name in enumerate(df.columns):
            worksheet.write(1, col_num, column_name, gray_format)
        
        # Adjust column widths
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, min(column_width, 30))
    
    output.seek(0)
    return output.getvalue()

def display_data_preview(df):
    """Display data with the two-row header only (no filters)"""
    
    # Display the two-row header
    col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 = st.columns(12)
    
    # First row - Orange headers
    with col1:
        st.markdown('<div class="orange-header">PORTFOLIO NAME</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="orange-header">SECURITY_ID</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="orange-header">SECURITY_ID</div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="orange-header">SECURITY_ID</div>', unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="orange-header">SECURITY_ID</div>', unsafe_allow_html=True)
    with col6:
        st.markdown('<div class="orange-header"></div>', unsafe_allow_html=True)  # Empty for Security Name
    with col7:
        st.markdown('<div class="orange-header">QUANTITY</div>', unsafe_allow_html=True)
    with col8:
        st.markdown('<div class="orange-header">"Fixed or<br>Drifting Weight"</div>', unsafe_allow_html=True)
    with col9:
        st.markdown('<div class="orange-header">Market Price</div>', unsafe_allow_html=True)
    with col10:
        st.markdown('<div class="orange-header">Cost Price</div>', unsafe_allow_html=True)
    with col11:
        st.markdown('<div class="orange-header">As of Date</div>', unsafe_allow_html=True)
    with col12:
        st.markdown('<div class="orange-header">Custom grouping</div>', unsafe_allow_html=True)
    
    # Second row - Gray headers (column names)
    col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 = st.columns(12)
    
    with col1:
        st.markdown('<div class="gray-header">Portfolio Name</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="gray-header">Security</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="gray-header">Sedol</div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="gray-header">Cusip</div>', unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="gray-header">ISIN</div>', unsafe_allow_html=True)
    with col6:
        st.markdown('<div class="gray-header">Security Name</div>', unsafe_allow_html=True)
    with col7:
        st.markdown('<div class="gray-header">Position</div>', unsafe_allow_html=True)
    with col8:
        st.markdown('<div class="gray-header">Weight</div>', unsafe_allow_html=True)
    with col9:
        st.markdown('<div class="gray-header">Mkt Px</div>', unsafe_allow_html=True)
    with col10:
        st.markdown('<div class="gray-header">Cost Price</div>', unsafe_allow_html=True)
    with col11:
        st.markdown('<div class="gray-header">As of Date</div>', unsafe_allow_html=True)
    with col12:
        st.markdown('<div class="gray-header">New Classification</div>', unsafe_allow_html=True)
    
    # Display only first 5 rows of data
    display_df = df.head(5)
    
    # Create a container for the data table
    table_container = st.container()
    with table_container:
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    st.caption(f"Showing first 5 of {len(df)} total rows")
    
    return df

def main():
    st.set_page_config(
        page_title="PD Capital Automation Tool",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Main header section
    st.markdown('<h1 class="main-title">PD CAPITAL AUTOMATION TOOL</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Upload Excel files and automate portfolio data processing with intelligent column mapping</p>', unsafe_allow_html=True)
    
    # Sidebar with comprehensive information
    with st.sidebar:
        st.markdown('<div class="sidebar-content">', unsafe_allow_html=True)
        
        st.markdown('<h3 style="color: #1E3A8A; margin-bottom: 20px;">📋 TOOL OVERVIEW</h3>', unsafe_allow_html=True)
        
        # How to Use section
        with st.expander("🚀 **HOW TO USE**", expanded=True):
            st.markdown("""
            1. **Upload** Excel files using the upload area
            2. **Process** files with one click
            3. **Review** output data preview
            4. **Download** combined Excel file
            
            **Quick Tip:** Select multiple files by holding Ctrl/Cmd
            """)
        
        # Required Columns section
        with st.expander("📋 **REQUIRED COLUMNS**", expanded=True):
            st.markdown("""
            Your input Excel files **must** contain:
            
            • `ISIN Code`
            • `Quantity / Notional`
            • `Cost Price (Ref. Cur)`
            • `Reporting Current Rate Date`
            
            ⚠️ Files missing these columns will be skipped
            """)
        
        # Column Mapping section
        with st.expander("🔄 **COLUMN MAPPING**", expanded=True):
            st.markdown("""
            **Automatic Mapping:**
            - `ISIN Code` → `ISIN` (E)
            - `Quantity / Notional` → `Position` (G)
            - `Cost Price (Ref. Cur)` → `Cost Price` (J)
            - `Reporting Current Rate Date` → `As of Date` (K)
            - Filename → `Portfolio Name` (A)
            
            **Empty Columns:**
            - Security (B)
            - Sedol (C)
            - Cusip (D)
            - Security Name (F)
            - Weight (H)
            - Mkt Px (I)
            - New Classification (L)
            """)
        
        # Limitations section
        with st.expander("⚠️ **LIMITATIONS**", expanded=True):
            st.markdown("""
            **Current Tool Limits:**
            
            • **ISIN Validation:** Only processes 12-character alphanumeric codes
            • **File Size:** Large files may take longer to process
            • **Column Names:** Must match exactly (case-sensitive)
            • **Excel Format:** Only .xlsx and .xls files supported
            • **Portfolio Name:** Limited to 10 characters from filename
            
            **Performance Notes:**
            - Processing time increases with file size
            - Multiple files processed sequentially
            - Output combines all files into single Excel
            """)
        
        # File Status section
        if 'processed_files' in st.session_state:
            with st.expander("📊 **PROCESSING STATUS**", expanded=True):
                st.markdown(f"""
                **Last Run:**
                • Processed: <span class="status-badge status-success">{st.session_state.processed_files} files</span>
                • Total Rows: {st.session_state.total_rows:,}
                • Valid ISINs: {st.session_state.valid_isins:,}
                """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Main content area
    st.markdown('<div class="section-header">📤 UPLOAD FILES</div>', unsafe_allow_html=True)
    
    # File upload area
    uploaded_files = st.file_uploader(
        "Drag and drop Excel files here or click to browse",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Select multiple Excel files for processing",
        label_visibility="collapsed"
    )
    
    if uploaded_files:
        # Show file selection status
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📁 Files Selected", len(uploaded_files))
        with col2:
            portfolio_names = [extract_portfolio_name(f.name) for f in uploaded_files]
            st.metric("🏷️ Unique Portfolios", len(set(portfolio_names)))
        with col3:
            # Estimate total rows
            estimated_rows = len(uploaded_files) * 100  # Rough estimate
            st.metric("📊 Estimated Rows", f"~{estimated_rows:,}")
        
        st.markdown("---")
        
        # Process files button
        if st.button("🚀 PROCESS ALL FILES", type="primary", use_container_width=True):
            with st.spinner("Processing your files..."):
                result = process_multiple_files(uploaded_files)
                
                if result:
                    df_combined, processed_files, failed_files = result
                    
                    if processed_files > 0:
                        # Store in session state for sidebar
                        st.session_state.processed_files = processed_files
                        st.session_state.total_rows = len(df_combined)
                        st.session_state.valid_isins = (df_combined['ISIN'] != "").sum()
                        
                        # Success message
                        st.markdown('<div class="success-box">✅ Successfully processed {} file(s)</div>'.format(processed_files), unsafe_allow_html=True)
                        
                        if failed_files:
                            st.markdown('<div class="warning-box">⚠️ Failed to process {} file(s): {}</div>'.format(
                                len(failed_files), ", ".join(failed_files[:3]) + ("..." if len(failed_files) > 3 else "")
                            ), unsafe_allow_html=True)
                        
                        st.markdown("---")
                        
                        # Output data preview
                        st.markdown('<div class="section-header">📊 OUTPUT DATA PREVIEW</div>', unsafe_allow_html=True)
                        st.markdown("*First 5 rows of the combined output data*")
                        
                        display_data_preview(df_combined)
                        
                        st.markdown("---")
                        
                        # Statistics section
                        st.markdown('<div class="section-header">📈 PROCESSING STATISTICS</div>', unsafe_allow_html=True)
                        
                        stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
                        
                        with stat_col1:
                            st.metric("Total Rows", f"{len(df_combined):,}")
                        with stat_col2:
                            valid_isins = (df_combined['ISIN'] != "").sum()
                            st.metric("Valid ISINs", f"{valid_isins:,}")
                        with stat_col3:
                            unique_portfolios = df_combined['Portfolio Name'].nunique()
                            st.metric("Unique Portfolios", unique_portfolios)
                        with stat_col4:
                            completion_rate = (processed_files / len(uploaded_files)) * 100
                            st.metric("Success Rate", f"{completion_rate:.1f}%")
                        
                        # Portfolio breakdown
                        if unique_portfolios > 1:
                            st.markdown("**Portfolio Breakdown:**")
                            portfolio_counts = df_combined['Portfolio Name'].value_counts()
                            for portfolio, count in portfolio_counts.items():
                                st.progress(min(count / len(df_combined), 1.0), 
                                          text=f"{portfolio}: {count} rows ({count/len(df_combined)*100:.1f}%)")
                        
                        st.markdown("---")
                        
                        # Download section
                        st.markdown('<div class="section-header">📥 DOWNLOAD OUTPUT</div>', unsafe_allow_html=True)
                        
                        try:
                            excel_data = to_excel(df_combined)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            portfolio_count = df_combined['Portfolio Name'].nunique()
                            output_filename = f"Portfolio_Data_{portfolio_count}Portfolios_{timestamp}.xlsx"
                            
                            # Download button with enhanced styling
                            col1, col2, col3 = st.columns([1, 2, 1])
                            with col2:
                                st.download_button(
                                    label="⬇️ DOWNLOAD EXCEL FILE",
                                    data=excel_data,
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    help=f"Download {len(df_combined):,} rows from {processed_files} files"
                                )
                            
                            # File info
                            st.markdown("""
                            <div class="info-card">
                            <strong>File Information:</strong><br>
                            • Format: Excel (.xlsx) with proper formatting<br>
                            • Header Style: Two-row headers (Orange + Gray)<br>
                            • Sheet Name: "Output"<br>
                            • Column Widths: Auto-adjusted for readability
                            </div>
                            """, unsafe_allow_html=True)
                            
                        except ImportError:
                            st.markdown("""
                            <div class="error-box">
                            ❌ **xlsxwriter not installed!**<br>
                            Please install xlsxwriter for Excel formatting:<br>
                            <code>pip install xlsxwriter</code><br>
                            Then restart the application.
                            </div>
                            """, unsafe_allow_html=True)
                        except Exception as e:
                            st.markdown(f'<div class="error-box">❌ Error creating Excel file: {str(e)}</div>', unsafe_allow_html=True)
                    
                    else:
                        st.markdown('<div class="error-box">❌ No files were successfully processed. Please check the column requirements.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-box">❌ Failed to process any files. Please check the uploaded files.</div>', unsafe_allow_html=True)
    
    else:
        # No files uploaded - show instructions
        st.markdown("""
        <div class="upload-area">
        <h3 style="color: #6B7280; margin-bottom: 15px;">📁 No files selected</h3>
        <p style="color: #9CA3AF; margin-bottom: 20px;">Drag and drop your Excel files here or click to browse</p>
        <p style="color: #6B7280; font-size: 14px;">Supported formats: .xlsx, .xls</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Quick start guide
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div class="info-card">
            <strong>💡 Getting Started</strong><br><br>
            1. Prepare your Excel files with required columns<br>
            2. Upload multiple files at once<br>
            3. Click "PROCESS ALL FILES"<br>
            4. Download the combined output
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="info-card">
            <strong>⚙️ Important Notes</strong><br><br>
            • Files are processed in sequence<br>
            • Each file gets a unique portfolio name<br>
            • Output combines all data into one file<br>
            • Empty columns remain empty as specified
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()