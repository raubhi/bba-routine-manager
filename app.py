import streamlit as st
import pandas as pd

st.set_page_config(page_title="BBA Routine Optimizer", layout="wide")
st.title("BBA Routine Management System")
st.markdown("### University of Scholars | Program Coordinator Dashboard")
st.divider()

st.subheader("1. Data Upload")
course_file = st.file_uploader("Upload Course Offering Sheet", type=["xlsx"])

if course_file:
    # Load all sheets to find the Faculty List
    xls = pd.ExcelFile(course_file)
    faculty_sheet = next((s for s in xls.sheet_names if 'faculty' in s.lower()), None)
    
    faculty_data = []
    if faculty_sheet:
        df_fac = pd.read_excel(xls, sheet_name=faculty_sheet)
        
        # Parser to detect the "Contractual" row and split employment types
        is_adjunct = False
        for index, row in df_fac.iterrows():
            first_col_val = str(row.iloc[0]).strip().lower()
            if first_col_val == 'contractual':
                is_adjunct = True
                continue
            
            name = str(row.iloc[1]).strip()
            if name.lower() not in ['nan', 'name', '']:
                emp_type = 'Adjunct' if is_adjunct else 'Full-Time'
                faculty_data.append({"Faculty Name": name, "Employment Type": emp_type})
    
    # Ensure TBA is always an option in the roster
    if not any(f['Faculty Name'] == 'TBA' for f in faculty_data):
        faculty_data.append({"Faculty Name": "TBA", "Employment Type": "Adjunct"})
        
    fac_df = pd.DataFrame(faculty_data)
    
    # Append the off-day matrix directly to the faculty list
    for day in ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]:
        fac_df[day] = False
        
    st.subheader("2. Faculty Roster & Off-Day Manager")
    st.info("Scroll to the bottom of the table and click the '+' to add a newly recruited faculty. Click any 'Employment Type' cell to toggle between Full-Time and Adjunct.")
    
    # The interactive editor handles toggles and adding rows automatically
    edited_faculty = st.data_editor(
        fac_df, 
        num_rows="dynamic",
        column_config={
            "Employment Type": st.column_config.SelectboxColumn(
                "Employment Type",
                options=["Full-Time", "Adjunct"],
                required=True
            )
        },
        use_container_width=True
    )
    
    st.divider()
    
    # --- TBA CONVERSION MANAGER ---
    st.subheader("3. TBA Conversion Manager")
    st.write("Assign newly added faculties to replace TBA slots before running the solver.")
    col_tba1, col_tba2 = st.columns(2)
    with col_tba1:
        replace_tba = st.toggle("Convert TBA slots for this generation?", value=False)
    with col_tba2:
        if replace_tba:
            available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
            tba_replacement = st.selectbox("Select Replacement Faculty:", available_faculties)

    st.divider()
    
    # --- GLOBAL SETTINGS ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Global Settings")
        include_majors = st.toggle("Include Major Courses", value=False)
    with col2:
        st.subheader("Senior Batch Scheduling")
        # Filter to only show Full-Time faculties for Sun/Mon selection
        ft_faculties = edited_faculty[edited_faculty['Employment Type'] == 'Full-Time']['Faculty Name'].tolist()
        sun_mon_faculties = st.multiselect("Select Faculties (Sun/Mon strictly):", ft_faculties)

    st.divider()
    
    if st.button("Generate Optimized Routine", type="primary", use_container_width=True):
        st.success("Routine generation triggered! The solver will now route TBA classes to your selected faculty.")
