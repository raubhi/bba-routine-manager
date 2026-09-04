import streamlit as st
import pandas as pd
from data_loader import process_uploads, export_routine
from solver import generate_routine

st.set_page_config(page_title="BBA Routine Optimizer", layout="wide")
st.title("BBA Routine Management System")
st.markdown("### University of Scholars | Program Coordinator Dashboard")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("Global Settings")
    include_majors = st.toggle("Include Major Courses", value=False)
    
with col2:
    st.subheader("Senior Batch Scheduling")
    sun_mon_faculties = st.multiselect(
        "Select Faculties (Sun/Mon strictly):",
        ["Dr. Rashed Chowdhury", "Dr. M. Sohel Rana", "Shahinur Rahman"] 
    )

st.divider()
st.subheader("Faculty Off-Day Manager")
blank_off_days = pd.DataFrame({
    "Faculty": ["Muntasir Hafij Nashek", "Tilottama Ahmed", "Nur-A-Alam Mishad"],
    "Saturday": [False, False, False], "Sunday": [False, False, False],
    "Monday": [False, False, False], "Tuesday": [False, False, False],
    "Wednesday": [False, False, False], "Thursday": [False, False, False]
})
edited_off_days = st.data_editor(blank_off_days, num_rows="dynamic", use_container_width=True)

st.divider()
st.subheader("Data Upload & Execution")
col3, col4 = st.columns(2)
with col3:
    course_file = st.file_uploader("1. Upload Course Offering", type=["xlsx", "csv"])
with col4:
    faculty_file = st.file_uploader("2. Upload Faculty Roster", type=["xlsx", "csv"])

if st.button("Generate Optimized Routine", type="primary", use_container_width=True):
    if course_file and faculty_file:
        with st.spinner("Processing data and running CP-SAT optimization..."):
            clean_data = process_uploads(course_file, faculty_file, include_majors)
            solved_schedule = generate_routine(clean_data, sun_mon_faculties, edited_off_days)
            excel_file = export_routine(solved_schedule)
            
            st.success("Routine successfully generated with zero clashes!")
            st.download_button(
                label="Download Excel Routine",
                data=excel_file,
                file_name="BBA_Fall_2026_Routine.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Please upload both templates to begin.")
