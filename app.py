import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="BBA Routine Optimizer", layout="wide")
st.title("BBA Routine Management System")
st.divider()

# Initialize Session State to store individual rules persistently
if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}

# --- 1. DATA UPLOAD & PARSING ---
st.subheader("1. Data Upload")
course_file = st.file_uploader("Upload Course Offering Sheet", type=["xlsx"])

if course_file:
    xls = pd.ExcelFile(course_file)
    faculty_sheet = next((s for s in xls.sheet_names if 'faculty' in s.lower()), None)
    
    faculty_data = []
    if faculty_sheet:
        df_fac = pd.read_excel(xls, sheet_name=faculty_sheet)
        is_adjunct = False
        for index, row in df_fac.iterrows():
            first_col = str(row.iloc[0]).strip().lower()
            if first_col == 'contractual':
                is_adjunct = True
                continue
            name = str(row.iloc[1]).strip()
            if name.lower() not in ['nan', 'name', '']:
                emp_type = 'Adjunct' if is_adjunct else 'Full-Time'
                faculty_data.append({"Faculty Name": name, "Type": emp_type})
    
    if not any(f['Faculty Name'] == 'TBA' for f in faculty_data):
        faculty_data.append({"Faculty Name": "TBA", "Type": "Adjunct"})
        
    fac_df = pd.DataFrame(faculty_data)
    days = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu"]
    for day in days:
        fac_df[f"{day} Max"] = 3
        
    st.subheader("2. Faculty Roster & Daily Limits")
    st.warning("⚠️ **Set a day to 0 to make it an Off-Day (Blocked).** Set to 1-6 to define maximum classes for that specific day.")
    
    col_config = {
        "Type": st.column_config.SelectboxColumn("Type", options=["Full-Time", "Adjunct"], required=True)
    }
    for day in days:
        col_config[f"{day} Max"] = st.column_config.NumberColumn(f"{day} Max", min_value=0, max_value=6, step=1)
    
    edited_faculty = st.data_editor(
        fac_df, 
        num_rows="dynamic",
        column_config=col_config,
        use_container_width=True
    )
    
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
    st.divider()

    # --- 3. ADVANCED CONSTRAINTS ENGINE (INDIVIDUAL RULES) ---
    st.subheader("3. Advanced Scheduling Constraints")
    col_adv1, col_adv2 = st.columns(2)
    
    with col_adv1:
        st.write("**Individual Faculty Slot Locks & Reservations**")
        st.info("Configure specific rules for one teacher at a time, then click 'Add Rule'.")
        
        selected_fac = st.selectbox("Select Faculty Member:", available_faculties)
        rule_day = st.selectbox("Select Day for this Teacher:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"])
        rule_slot = st.selectbox("Select Time Slot for this Teacher:", ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"])
        
        if st.button("➕ Add Faculty Rule"):
            st.session_state.fac_rules.append({"Faculty": selected_fac, "Day": rule_day, "Slot": rule_slot})
            st.success(f"Added rule for {selected_fac} on {rule_day} at {rule_slot}")
            
        if st.session_state.fac_rules:
            st.write("Active Faculty Rules:")
            st.table(pd.DataFrame(st.session_state.fac_rules))
            if st.button("Clear Faculty Rules"):
                st.session_state.fac_rules = []
                st.rerun()
            
    with col_adv2:
        st.write("**Batch Active Day Windows & Blackouts**")
        st.info("Configure day limits and grey out unavailable days per batch.")
        
        selected_batch = st.selectbox("Select Batch:", ["20th", "19th", "18th", "17th", "16th", "15th", "14th"])
        batch_max_days = st.number_input(f"Max Active Days for {selected_batch}:", min_value=1, max_value=6, value=3)
        blocked_batch_days = st.multiselect(f"Grey Out / Blocked Days for {selected_batch}:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"])
        
        if st.button("💾 Save Batch Configuration"):
            st.session_state.batch_rules[selected_batch] = {"Max Days": batch_max_days, "Blocked Days": blocked_batch_days}
            st.success(f"Saved configuration for {selected_batch}")
            
        if st.session_state.batch_rules:
            st.write("Saved Batch Configurations:")
            st.json(st.session_state.batch_rules)
            
        st.write("**TBA Conversion Manager**")
        replace_tba = st.toggle("Convert TBA slots for this generation?", value=False)
        tba_replacement = None
        if replace_tba:
            tba_replacement = st.selectbox("Replacement Faculty:", available_faculties)

    st.divider()

    # --- 4. GLOBAL SETTINGS ---
    st.subheader("4. Global Engine Settings")
    g_col1, g_col2 = st.columns(2)
    with g_col1:
        include_majors = st.toggle("Include Major Courses", value=False)
    with g_col2:
        st.toggle("Enforce Compact Scheduling (Prevent extreme gaps)", value=True)

    st.divider()
    
    # --- 5. EXECUTION & IN-APP PREVIEW WINDOW ---
    if st.button("Generate & Preview Routine", type="primary", use_container_width=True):
        with st.spinner("Running optimization engine..."):
            
            # Mocking solver output structure for previewing
            preview_data = [
                {"Batch": "20th Batch (Section A)", "Day": "Saturday", "Time Slot": "9:30-11:00", "Course Code": "BBA 1101-0413", "Course Title": "Introduction To Business", "Faculty": "Muntasir Hafij Nashek"},
                {"Batch": "20th Batch (Section B)", "Day": "Saturday", "Time Slot": "11:00-12:30", "Course Code": "ECO 1103-0311", "Course Title": "Microeconomics", "Faculty": "Tilottama Ahmed"},
                {"Batch": "19th Batch (Section A)", "Day": "Sunday", "Time Slot": "9:30-11:00", "Course Code": "BBA 1201-0413", "Course Title": "Principles of Management", "Faculty": "Shadia Sharmin"},
                {"Batch": "18th Batch (Section A)", "Day": "Tuesday", "Time Slot": "1:30-3:00", "Course Code": "BBA 2101-0414", "Course Title": "Bank Management", "Faculty": "Dr. Md Abu Hasnat"}
            ]
            df_preview = pd.DataFrame(preview_data)
            
            # Store in session state so it stays visible on screen
            st.session_state.preview_df = df_preview

    # Display Preview Window if data exists in session state
    if 'preview_df' in st.session_state:
        st.subheader("📊 Live Routine Preview Window")
        st.info("Review the generated routine below. If adjustments are needed, modify your rules above and re-generate. Once satisfied, download the official Excel template format.")
        
        st.dataframe(st.session_state.preview_df, use_container_width=True)
        
        # Build Excel openpyxl file for final download
        wb = Workbook()
        ws = wb.active
        ws.title = "Fall 2026 Routine"
        
        # Writing data into the template layout format
        sample_routine = [
            {'code': 'BBA 1101-0413', 'title': 'Introduction To Business', 'faculty': 'Muntasir Hafij Nashek', 'row': 2, 'col': 2},
            {'code': 'ECO 1103-0311', 'title': 'Microeconomics', 'faculty': 'Tilottama Ahmed', 'row': 2, 'col': 3},
            {'code': 'BBA 1104-0414', 'title': 'Principles of Marketing', 'faculty': 'Nur-A-Alam Mishad', 'row': 3, 'col': 2}
        ]
        
        for entry in sample_routine:
            cell_value = f"{entry['code']}\n{entry['title']}\n{entry['faculty']}"
            cell = ws.cell(row=entry['row'], column=entry['col'])
            cell.value = cell_value
            cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            cell.font = Font(name='Arial', size=10)
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Finalized Excel Routine (.xlsx)",
            data=output,
            file_name="BBA_Fall_2026_Routine_Optimized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
