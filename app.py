import streamlit as st
import pandas as pd
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
import os
import tempfile
from io import BytesIO

# Custom CSS for modern dark mode design
st.markdown("""
<style>
    /* General modern styling with dark mode */
    .stApp {
        background-color: #121212;
        color: #EDEDED;
    }
    h1, h2, h3, .stMarkdown, .stText {
        color: #EDEDED;
    }
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #1F1F1F;
        border-right: 1px solid #333333;
    }
    [data-testid="stSidebar"] .stButton > button {
        background-color: #2196F3;
        color: #FFFFFF;
        border-radius: 8px;
    }
    /* Expander styling */
    .stExpander {
        border: 1px solid #333333;
        border-radius: 8px;
        background-color: #1F1F1F;
    }
    .stExpander summary {
        color: #EDEDED;
        font-weight: bold;
    }
    /* Button styling */
    .stButton > button {
        background-color: #4CAF50;
        color: #FFFFFF;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 16px;
        transition: background-color 0.3s;
    }
    .stButton > button:hover {
        background-color: #388E3C;
    }
    /* Download button */
    .stDownloadButton > button {
        background-color: #2196F3;
        color: #FFFFFF;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 16px;
        transition: background-color 0.3s;
    }
    .stDownloadButton > button:hover {
        background-color: #1976D2;
    }
    /* Dataframe styling */
    .stDataFrame {
        border: 1px solid #333333;
        border-radius: 8px;
        background-color: #1F1F1F;
    }
    .stDataFrame th, .stDataFrame td {
        color: #EDEDED;
    }
    /* Slider styling */
    .stSlider .stMarkdown {
        color: #EDEDED;
    }
    /* Progress bar */
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    /* Info, success, warning */
    .stAlert {
        border-radius: 8px;
    }
    .stInfo {
        background-color: #1F1F1F;
        color: #BBDEFB;
    }
    .stSuccess {
        background-color: #1F1F1F;
        color: #A5D6A7;
    }
    .stWarning {
        background-color: #1F1F1F;
        color: #FFCC80;
    }
</style>
""", unsafe_allow_html=True)

# Function to download image
def download_image(url, temp_dir):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        file_path = os.path.join(temp_dir, os.path.basename(url.split('?')[0]))
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return file_path
    except Exception as e:
        st.warning(f"Error downloading {url}: {e}")
        return None

# Function to convert CSV to XLSX
def csv_to_xlsx(file_bytes, temp_dir):
    df = pd.read_csv(BytesIO(file_bytes))
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for col_idx, column in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx).value = column
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx).value = value
    temp_xlsx = os.path.join(temp_dir, "converted.xlsx")
    wb.save(temp_xlsx)
    return temp_xlsx, df

# Sidebar for navigation and settings
with st.sidebar:
    st.title("App Settings")
    st.markdown("Customize sizes and controls")
    
    # User controls for sizes
    img_width = st.slider("Image Width (pixels)", min_value=50, max_value=300, value=100, step=10)
    img_height = st.slider("Image Height (pixels)", min_value=50, max_value=300, value=100, step=10)
    row_height = st.slider("Row Height (points)", min_value=40, max_value=200, value=80, step=10)
    col_width = st.slider("Column Width (characters)", min_value=10, max_value=50, value=20, step=1)
    
    st.markdown("---")
    st.caption("Modern Image Embedder App v1.0")
    st.caption("Made with ❤️ by Hassanelb")

# Main content with collapsible expanders
st.title("Modern CSV/Excel Image Embedder")
st.markdown("Enhance your spreadsheets by embedding images from URLs. Use the sections below to proceed.")

# Expander for File Upload
with st.expander("Step 1: Upload File", expanded=True):
    uploaded_file = st.file_uploader("Choose a CSV or Excel file", type=["csv", "xlsx"], help="Upload your file here.")

# Process if file is uploaded
if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    
    # Determine file type and load sheet names
    if uploaded_file.name.endswith('.csv'):
        sheet_names = ["Sheet1"]
        df = pd.read_csv(BytesIO(file_bytes))
    else:
        xl = pd.ExcelFile(BytesIO(file_bytes))
        sheet_names = xl.sheet_names
        df = None
    
    # Expander for Sheet Selection
    with st.expander("Step 2: Select Sheet", expanded=True):
        sheet_name = st.selectbox("Select the sheet", sheet_names, help="Choose the sheet to process.")
    
    if sheet_name:
        # Load preview data
        if uploaded_file.name.endswith('.csv'):
            preview_df = df
        else:
            preview_df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
        
        # Expander for File Preview
        with st.expander("Step 3: File Preview", expanded=True):
            st.dataframe(preview_df.head(10), use_container_width=True)
        
        # Expander for Column Selection
        with st.expander("Step 4: Select Column", expanded=True):
            column_headers = preview_df.columns.tolist()
            image_col_name = st.selectbox("Select the column with image URLs", column_headers, help="Choose the column containing image URLs.")
        
        # Expander for Insert Images
        with st.expander("Step 5: Insert Images", expanded=True):
            if st.button("Start Processing", use_container_width=True):
                with st.spinner("Processing images... Please wait."):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        # Handle CSV or XLSX
                        if uploaded_file.name.endswith('.csv'):
                            temp_xlsx, _ = csv_to_xlsx(file_bytes, temp_dir)
                        else:
                            temp_xlsx = os.path.join(temp_dir, "uploaded.xlsx")
                            with open(temp_xlsx, "wb") as f:
                                f.write(file_bytes)
                        
                        # Load workbook
                        wb = load_workbook(temp_xlsx)
                        ws = wb[sheet_name]
                        
                        # Insert new column
                        image_col = column_headers.index(image_col_name) + 1
                        new_image_col = image_col + 1
                        ws.insert_cols(new_image_col)
                        ws.cell(row=1, column=new_image_col).value = "Embedded Image"
                        
                        # Process images with progress bar
                        num_rows = ws.max_row - 1
                        progress_bar = st.progress(0.0)
                        processed = 0
                        
                        for row in range(2, ws.max_row + 1):
                            url_cell = ws.cell(row=row, column=image_col)
                            url = url_cell.value
                            if url:
                                image_path = download_image(url, temp_dir)
                                if image_path:
                                    img = Image(image_path)
                                    img.width = img_width
                                    img.height = img_height
                                    ws.add_image(img, ws.cell(row=row, column=new_image_col).coordinate)
                                    ws.row_dimensions[row].height = row_height
                                else:
                                    ws.cell(row=row, column=new_image_col).value = "Failed to download"
                            
                            processed += 1
                            progress_bar.progress(processed / num_rows)
                        
                        # Adjust column width
                        ws.column_dimensions[ws.cell(row=1, column=new_image_col).column_letter].width = col_width
                        
                        # Save to BytesIO
                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                
                # Success message and download
                st.success("Processing complete! Download the modified file.")
                st.download_button(
                    label="Download Modified Excel",
                    data=output,
                    file_name="data_with_images.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
else:
    st.info("Upload a file to begin.")