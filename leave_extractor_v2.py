import streamlit as st
import google.generativeai as genai
from PIL import Image
import json
import pandas as pd
import io

# ==========================================
# 1. CONFIGURATION
# ==========================================
API_KEY = st.secrets["Gen_API"]["API_KEY"] # Replace with your actual key
genai.configure(api_key=API_KEY)

# Initialize both models based on your previous code
model_flash = genai.GenerativeModel('gemini-3-flash-preview')
model_pro = genai.GenerativeModel('gemini-1.5-pro')

st.set_page_config(page_title='Zenxin Document Extractor', layout='wide')

# ==========================================
# 2. SIDEBAR - MODE SELECTION
# ==========================================
st.sidebar.title("Settings")
st.sidebar.write("Select Mode")
app_mode = st.sidebar.radio(
    label="Select Mode",
    label_visibility="collapsed", # Hides the duplicate label to match your image
    options=("Attendance Report OCR", "Leave/Overtime OCR")
)

# ==========================================
# 3. MAIN UI & IMAGE INPUT
# ==========================================
st.title(f'📄 {app_mode}')
st.write('Upload or capture a photo of the document.')

upload_option = st.radio('Input Method:', ('Camera', 'Upload Image'), horizontal=True)

if upload_option == 'Camera':
    # Streamlit's camera_input only handles one image at a time natively
    camera_file = st.camera_input('Capture Report')
    uploaded_files = [camera_file] if camera_file else []
else:
    # File uploader configured to accept multiple images
    uploaded_files = st.file_uploader(
        'Upload Report Pages', 
        type=['jpg', 'png', 'jpeg'], 
        accept_multiple_files=True
    )

# ==========================================
# 4. ANALYSIS & DISPLAY
# ==========================================
if uploaded_files:
    # Convert uploaded files to PIL Images
    images = [Image.open(file) for file in uploaded_files]
    
    col1, col2 = st.columns([1, 1.5])
    
    with col1:
        st.image(images, use_container_width=True)
    
    with col2:
        # ------------------------------------------
        # MODE A: ATTENDANCE REPORT
        # ------------------------------------------
        if app_mode == "Attendance Report OCR":
            if st.button('Extract Data & Generate Excel', use_container_width=True):
                with st.spinner('Analyzing handwriting and formatting for Excel...'):
                    prompt = '''
                    You are an expert at transcribing messy handwritten notes on printed timesheets. 
                    Extract the data and return ONLY a valid JSON object. Do NOT include markdown formatting like ```json in the output.

                    ### 🚨 HANDWRITING DICTIONARY & RULES 🚨
                    1. Known Codes: 'UPL', 'UPH', 'MC', 'AL', 'OT', 'OFF'.
                    2. Column Locations: Handwriting is usually over 'Resume', 'Out', and 'OT'.
                    3. Circled Numbers: Transcribe as "UPL - 1 (Circled [Number])".
                    4. Crossed-out Text: If a number is crossed out, include it like "~~5.5 Hours~~".
                    5. Checkmarks: Note any checkmarks (✓).
                    6. Margin Notes: Look for Malay text like "Lambat masuk" or math.

                    Return EXACTLY this JSON structure:
                    {
                      "employee": {
                        "User ID": "String", "Name": "String", "Department": "String", "Date Range": "String"
                      },
                      "timesheet": [
                        {"Date": "String", "Day": "String", "In": "String", "Out": "String", "Work Hrs": "String", "Remarks": "String"}
                      ],
                      "summary": [
                        "String"
                      ],
                      "notes": [
                        "String"
                      ]
                    }
                    '''
                    
                    try:
                        request_content = [prompt] + images # <--- This flattens them into one list

                        response = model_flash.generate_content(
                            request_content, 
                            generation_config=genai.types.GenerationConfig(temperature=0.1)
                        )
                        
                        clean_json = response.text.replace('```json', '').replace('```', '').strip()
                        data = json.loads(clean_json)

                        # --- ADD THIS SAFEGUARD ---
                        # If Gemini wraps the JSON in a list, extract the first dictionary
                        if isinstance(data, list):
                            data = data[0]
                        # --------------------------

                        st.success('Extraction Complete!')
                        
                        st.markdown("### Employee Information")
                        for key, value in data.get('employee', {}).items():
                            st.markdown(f"* **{key}:** {value}")
                            
                        st.markdown("### Timesheet Data")
                        df_timesheet = pd.DataFrame(data.get('timesheet', []))
                        st.dataframe(df_timesheet, use_container_width=True)
                        
                        # --- EXCEL GENERATION ---
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_timesheet.to_excel(writer, index=False, sheet_name='Timesheet Data')
                            df_emp = pd.DataFrame(list(data.get('employee', {}).items()), columns=['Field', 'Value'])
                            df_emp.to_excel(writer, index=False, sheet_name='Employee Info')
                            
                        excel_data = output.getvalue()
                        
                        st.download_button(
                            label="📥 Download as Excel (.xlsx)",
                            data=excel_data,
                            file_name="timesheet_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                        
                    except Exception as e:
                        st.error(f'Error parsing data. Please try again.')
                        st.expander("View Error Log").write(e)

        # ------------------------------------------
        # MODE B: LEAVE/OVERTIME RECORD
        # ------------------------------------------
        elif app_mode == "Leave/Overtime OCR":
            if st.button('Extract Overtime Data & Generate Excel', use_container_width=True):
                with st.spinner('Analyzing all OT records and calculating hours...'):
                    # Changed structure to an ARRAY [] so it can hold multiple employees/pages
                    prompt = """
                    You are a data extraction assistant. Analyze all the provided Leave/Overtime/Replacement Record images.
                    Extract the information and format it EXACTLY as the JSON structure below. 
                    Calculate the totals accurately. Return ONLY valid JSON, do NOT include markdown formatting like ```json.

                    ### 🚨 EXTRACTION RULES 🚨
                    1. Extract all rows from the table across ALL provided images. Empty cells should be an empty string "".
                    2. Convert fractions in hours to decimals (e.g., "1 ½" becomes 1.5).
                    3. Calculate the exact count of each record type (e.g., how many days of OT, AL, UPL).
                    4. Calculate the total sum of OT hours.

                    Return EXACTLY this JSON structure as a LIST (Array) of objects, one object per document/employee:
                    [
                      {
                        "employee": {
                          "Company": "String", "Name": "String", "Position": "String", "Month": "String", "Emp No": "String"
                        },
                        "records": [
                          {"Date": "String", "Type (Leave/OT)": "String", "Reason/Remark": "String", "Time From": "String", "Time To": "String", "Hours": "Number"}
                        ],
                        "summary": {
                          "Total OT Hours": "Number",
                          "Total OT Days": "Number",
                          "Total AL Days": "Number",
                          "Total UPL Days": "Number",
                          "Total MC Days": "Number"
                        },
                        "signatures": {
                          "Confirmed By": "String",
                          "Approval By": "String"
                        }
                      }
                    ]
                    """

                    try:
                        request_content = [prompt] + images
                        # Using model_flash based on your setup, but model_pro is recommended for multi-page tables
                        response = model_flash.generate_content(
                            request_content,
                            generation_config=genai.types.GenerationConfig(temperature=0.1)
                        )
                        
                        clean_json = response.text.replace('```json', '').replace('```', '').strip()
                        data = json.loads(clean_json)

                        # SAFEGUARD: Ensure data is always a list
                        if not isinstance(data, list):
                            data = [data]

                        st.success(f'Extraction Complete! Processed {len(data)} records.')
                        
                        # --- PREPARE MASTER LISTS FOR EXCEL ---
                        all_records_df = []
                        all_summaries_df = []

                        # --- 1. BUILD THE UI (Loop through each extracted sheet) ---
                        for idx, sheet in enumerate(data):
                            with st.expander(f"📄 Record {idx + 1}: {sheet.get('employee', {}).get('Name', 'Unknown')}", expanded=True):
                                st.markdown("### Employee Information")
                                for key, value in sheet.get('employee', {}).items():
                                    st.markdown(f"* **{key}:** {value}")
                                    
                                st.markdown("### Overtime & Leave Records")
                                df_records = pd.DataFrame(sheet.get('records', []))
                                st.dataframe(df_records, use_container_width=True)
                                
                                col_sum1, col_sum2 = st.columns(2)
                                with col_sum1:
                                    st.markdown("### Leave/OT Summary")
                                    for key, value in sheet.get('summary', {}).items():
                                        st.markdown(f"* **{key}:** {value}")
                                with col_sum2:
                                    st.markdown("### Signatures")
                                    for key, value in sheet.get('signatures', {}).items():
                                        st.markdown(f"* **{key}:** {value}")

                            # --- COMPILE DATA FOR MASTER EXCEL ---
                            emp_name = sheet.get('employee', {}).get('Name', 'Unknown')
                            emp_no = sheet.get('employee', {}).get('Emp No', 'Unknown')
                            month = sheet.get('employee', {}).get('Month', 'Unknown')

                            # Inject Employee details into the records so we know whose OT is whose in the master Excel
                            if not df_records.empty:
                                df_records.insert(0, 'Month', month)
                                df_records.insert(0, 'Emp No', emp_no)
                                df_records.insert(0, 'Employee Name', emp_name)
                                all_records_df.append(df_records)

                            # Create a flat row for the summary sheet
                            summary_row = sheet.get('employee', {})
                            summary_row.update(sheet.get('summary', {}))
                            all_summaries_df.append(summary_row)

                        # --- 2. GENERATE MASTER EXCEL FILE ---
                        if all_records_df:
                            final_records = pd.concat(all_records_df, ignore_index=True)
                            final_summaries = pd.DataFrame(all_summaries_df)

                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                # Write combined master sheets
                                final_records.to_excel(writer, index=False, sheet_name='All OT Records')
                                final_summaries.to_excel(writer, index=False, sheet_name='All Summaries')
                                
                            excel_data = output.getvalue()
                            
                            st.download_button(
                                label="📥 Download Master Excel (.xlsx)",
                                data=excel_data,
                                file_name="Master_Leave_OT_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )
                        
                    except Exception as e:
                        st.error(f'Error parsing data. Please try again.')
                        st.expander("View Error Log").write(e)