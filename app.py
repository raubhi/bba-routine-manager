import streamlit as st
import pandas as pd

st.set_page_config(page_title="BBA Routine Optimizer", layout="wide")
st.title("BBA Routine Management System")
st.markdown("### University of Scholars | Program Coordinator Dashboard")
st.divider()

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
    
    # Clarified naming: "Block [Day]" means off-day. Max Classes defaults to 3.
    for day in ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu"]:
        fac_df[f"Block {day}"] = False
    fac_df["Max Classes/Day"] = 3
        
    st.subheader("2. Faculty Roster & Base Constraints")
    st.warning("⚠️ **CHECKMARK (☑) = OFF-DAY (BLOCKED).** The faculty will NOT be assigned classes on checked days.")
    
    edited_faculty = st.data_editor(
        fac_df, 
        num_rows="dynamic",
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Full-Time", "Adjunct"], required=True),
            "Max Classes/Day": st.column_config.NumberColumn("Max Classes/Day", min_value=1, max_value=6, step=1)
        },
        use_container_width=True
    )
    
    # Extract clean list of available faculties for dropdowns
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()

    st.divider()

    # --- 3. ADVANCED CONSTRAINTS ENGINE ---
    st.subheader("3. Advanced Scheduling Constraints")
    
    col_adv1, col_adv2 = st.columns(2)
    with col_adv1:
        st.write("**Faculty Slot Reservations & Locks**")
        st.info("Force specific faculty members into specific days and time slots.")
        target_fac = st.selectbox("Select Faculty to Lock:", ["None"] + available_faculties)
        if target_fac != "None":
            lock_day = st.selectbox("Lock to Day:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"])
            lock_slot = st.selectbox("Lock to Slot:", ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"])
            st.button(f"Save Lock: {target_fac} on {lock_day} at {lock_slot}")
            
        st.write("**Faculty Maximums on Specific Days**")
        max_fac = st.selectbox("Select Faculty:", ["None"] + available_faculties, key="max_fac")
        if max_fac != "None":
            max_day = st.selectbox("Select Day:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"], key="max_fac_day")
            max_limit = st.number_input(f"Max classes for {max_fac} on {max_day}:", min_value=1, max_value=4, value=2)
            
    with col_adv2:
        st.write("**Batch Overrides**")
        st.info("Force a specific batch to have maximum classes on a specific day.")
        target_batch = st.selectbox("Select Batch:", ["None", "20th", "19th", "18th", "17th", "16th", "15th", "14th"])
        if target_batch != "None":
            batch_day = st.selectbox("Target Day:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"], key="b_day")
            batch_max = st.number_input(f"Max classes for {target_batch} on {batch_day}:", min_value=1, max_value=4, value=3)
            
        st.write("**TBA Conversion**")
        replace_tba = st.toggle("Convert TBA slots for this generation?", value=False)
        if replace_tba:
            tba_replacement = st.selectbox("Select Replacement Faculty:", available_faculties)

    st.divider()

    # --- 4. GLOBAL OPTIMIZATION SETTINGS ---
    st.subheader("4. Global Engine Settings")
    g_col1, g_col2 = st.columns(2)
    with g_col1:
        include_majors = st.toggle("Include Major Courses", value=False)
    with g_col2:
        compact_schedule = st.toggle("Enforce Compact Scheduling (Prevent extreme gaps for faculty)", value=True)

    st.divider()
    
    if st.button("Generate Optimized Routine", type="primary", use_container_width=True):
        st.success("Configuration captured! Backend solver integration is ready for the next phase.")
