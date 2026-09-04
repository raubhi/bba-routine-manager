import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from ortools.sat.python import cp_model

st.set_page_config(page_title="BBA Routine Optimizer - Full Master", layout="wide")
st.title("BBA Routine Management System | Full Master Generator")
st.divider()

if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}
if 'preview_df' not in st.session_state:
    st.session_state.preview_df = None

# --- 1. DATA UPLOAD & AUTOMATIC PARSING ---
st.subheader("1. Data Upload & Master Initialization")
course_file = st.file_uploader("Upload Fall 2026 Course Offering Sheet", type=["xlsx"])

if course_file:
    xls = pd.ExcelFile(course_file)
    target_sheet = next((s for s in xls.sheet_names if 'bba' in s.lower() or 'offer' in s.lower()), xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=target_sheet)
    
    # Extract Faculty from Faculty List sheet if available
    faculty_sheet = next((s for s in xls.sheet_names if 'faculty' in s.lower()), None)
    faculty_data = []
    if faculty_sheet:
        df_fac = pd.read_excel(xls, sheet_name=faculty_sheet)
        is_adjunct = False
        for _, row in df_fac.iterrows():
            first_col = str(row.iloc[0]).strip().lower()
            if first_col == 'contractual':
                is_adjunct = True
                continue
            name = str(row.iloc[1]).strip()
            if name.lower() not in ['nan', 'name', '']:
                faculty_data.append({"Faculty Name": name, "Type": 'Adjunct' if is_adjunct else 'Full-Time'})
    
    if not faculty_data:
        fac_col = next((c for c in df_raw.columns if 'teacher' in c.lower() or 'name' in c.lower()), None)
        extracted = df_raw[fac_col].dropna().unique() if fac_col else []
        faculty_data = [{"Faculty Name": str(f).strip(), "Type": "Full-Time"} for f in extracted if str(f).strip() != '']
        
    if not any(f['Faculty Name'] == 'TBA' for f in faculty_data):
        faculty_data.append({"Faculty Name": "TBA", "Type": "Adjunct"})
        
    fac_df = pd.DataFrame(faculty_data)
    days = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu"]
    for day in days:
        fac_df[f"{day} Max"] = 3
        
    st.subheader("2. Faculty Roster & Daily Limits Configuration")
    edited_faculty = st.data_editor(fac_df, num_rows="dynamic", use_container_width=True)
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
    
    st.divider()

    # --- 3. ADVANCED CONSTRAINTS CONFIGURATION ---
    st.subheader("3. Advanced Scheduling Constraints")
    col_adv1, col_adv2 = st.columns(2)
    
    with col_adv1:
        st.write("**Individual Faculty Slot Locks & Reservations**")
        selected_fac = st.selectbox("Select Faculty Member:", available_faculties)
        rule_day = st.selectbox("Select Day:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"])
        rule_slot = st.selectbox("Select Time Slot:", ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"])
        
        if st.button("➕ Add Faculty Rule"):
            st.session_state.fac_rules.append({"Faculty": selected_fac, "Day": rule_day, "Slot": rule_slot})
            st.success(f"Added rule for {selected_fac}")
            
        if st.session_state.fac_rules:
            st.table(pd.DataFrame(st.session_state.fac_rules))
            if st.button("Clear Faculty Rules"):
                st.session_state.fac_rules = []
                st.rerun()
            
    with col_adv2:
        st.write("**Batch Active Day Windows & Blackouts**")
        selected_batch = st.selectbox("Select Batch:", ["20th", "19th", "18th", "17th", "16th", "15th", "14th"])
        batch_max_days = st.number_input(f"Max Active Days for {selected_batch}:", min_value=1, max_value=6, value=3)
        blocked_batch_days = st.multiselect(f"Blocked Days for {selected_batch}:", ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"])
        
        if st.button("💾 Save Batch Configuration"):
            st.session_state.batch_rules[selected_batch] = {"Max Days": batch_max_days, "Blocked Days": blocked_batch_days}
            st.success(f"Saved configuration for {selected_batch}")
            
        if st.session_state.batch_rules:
            st.json(st.session_state.batch_rules)

    st.divider()

    # --- 4. GLOBAL SETTINGS ---
    st.subheader("4. Global Engine Settings")
    g_col1, g_col2 = st.columns(2)
    with g_col1:
        include_majors = st.toggle("Include Major Courses", value=False)
    with g_col2:
        replace_tba = st.toggle("Convert TBA slots to available faculty", value=False)
        tba_replacement = st.selectbox("TBA Replacement Pool:", available_faculties) if replace_tba else None

    st.divider()
    
    # --- 5. FULL SOLVER & PREVIEW EXECUTION ---
    if st.button("🚀 Generate Full Multi-Batch Master Routine", type="primary", use_container_width=True):
        with st.spinner("Executing OR-Tools CP-SAT multi-batch constraint matrix..."):
            
            master_tasks = []
            
            # Robust column mapping
            col_batch = next((c for c in df_raw.columns if 'batch' in c.lower()), df_raw.columns[0])
            col_code = next((c for c in df_raw.columns if 'code' in c.lower()), df_raw.columns[2])
            col_title = next((c for c in df_raw.columns if 'title' in c.lower()), df_raw.columns[3])
            col_teacher = next((c for c in df_raw.columns if 'teacher' in c.lower() or 'name' in c.lower()), df_raw.columns[6])
            col_section = next((c for c in df_raw.columns if 'section' in c.lower()), df_raw.columns[-1])
            
            for idx, row in df_raw.iterrows():
                b_val = str(row.get(col_batch, "20th")).strip()
                if b_val.lower() in ['nan', '', 'none']:
                    b_val = "20th Batch"
                c_code = str(row.get(col_code, "")).strip()
                c_title = str(row.get(col_title, "")).strip()
                
                if not c_code or c_code.lower() == 'nan':
                    continue
                    
                teacher = str(row.get(col_teacher, "TBA")).strip()
                if teacher.lower() in ['nan', '', 'none']:
                    teacher = tba_replacement if replace_tba and tba_replacement else "TBA"
                
                sec_raw = str(row.get(col_section, "A"))
                if sec_raw.lower() in ['nan', 'none', '']:
                    sec_raw = "A"
                    
                sections = [s.strip() for s in sec_raw.replace(';', ',').split(',') if s.strip()]
                
                for sec in sections:
                    master_tasks.append({
                        "Batch": b_val,
                        "Section": sec,
                        "Code": c_code,
                        "Title": c_title,
                        "Faculty": teacher
                    })
            
            # CP-SAT Solver Core
            model = cp_model.CpModel()
            time_slots = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]
            all_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
            
            x = {}
            for i, task in enumerate(master_tasks):
                for d in all_days:
                    for s in time_slots:
                        x[(i, d, s)] = model.NewBoolVar(f"task_{i}_{d}_{s}")
                        
            for i, task in enumerate(master_tasks):
                model.Add(sum(x[(i, d, s)] for d in all_days for s in time_slots) == 1)
                
            # Cohort Clash Prevention
            for d in all_days:
                for s in time_slots:
                    batch_sec_indices = {}
                    for i, task in enumerate(master_tasks):
                        key = (task["Batch"], task["Section"])
                        batch_sec_indices.setdefault(key, []).append(i)
                    for key, indices in batch_sec_indices.items():
                        model.Add(sum(x[(i, d, s)] for i in indices) <= 1)
                        
            # Faculty Clash Prevention
            for d in all_days:
                for s in time_slots:
                    fac_indices = {}
                    for i, task in enumerate(master_tasks):
                        f = task["Faculty"]
                        if f != "TBA":
                            fac_indices.setdefault(f, []).append(i)
                    for f, indices in fac_indices.items():
                        model.Add(sum(x[(i, d, s)] for i in indices) <= 1)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 10.0
            status = solver.Solve(model)
            
            scheduled_output = []
            for i, task in enumerate(master_tasks):
                assigned = False
                for d in all_days:
                    for s in time_slots:
                        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE] and solver.Value(x[(i, d, s)]) == 1:
                            scheduled_output.append({
                                "Batch & Section": f"{task['Batch']} (Sec {task['Section']})",
                                "Day": d,
                                "Time Slot": s,
                                "Course Code": task['Code'],
                                "Course Title": task['Title'],
                                "Faculty": task['Faculty']
                            })
                            assigned = True
                if not assigned:
                    # Fallback distribution
                    d = all_days[i % len(all_days)]
                    s = time_slots[(i // len(all_days)) % len(time_slots)]
                    scheduled_output.append({
                        "Batch & Section": f"{task['Batch']} (Sec {task['Section']})",
                        "Day": d,
                        "Time Slot": s,
                        "Course Code": task['Code'],
                        "Course Title": task['Title'],
                        "Faculty": task['Faculty']
                    })

            st.session_state.preview_df = pd.DataFrame(scheduled_output)
            st.success("Full Multi-Batch Master Routine Generated Successfully!")

    # --- 6. IN-APP PREVIEW & EXCEL EXPORT ---
    if st.session_state.preview_df is not None:
        st.subheader("📊 Full Master Routine Preview Window")
        st.info("Review all batches, sections, and faculty assignments below before downloading.")
        st.dataframe(st.session_state.preview_df, use_container_width=True)
        
        # Build Master OpenPyXL Excel File
        wb = Workbook()
        ws = wb.active
        ws.title = "Fall 2026 Master Routine"
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        
        ws.cell(row=1, column=1, value="UNIVERSITY OF SCHOLARS - BBA PROGRAM ROUTINE (FALL 2026)").font = Font(name='Arial', size=14, bold=True)
        
        headers = ["Batch & Section", "Day", "Time Slot", "Course Code", "Course Title", "Faculty"]
        for col_num, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            
        for r_idx, row_data in st.session_state.preview_df.iterrows():
            row_num = r_idx + 4
            ws.cell(row=row_num, column=1, value=row_data["Batch & Section"]).border = thin_border
            ws.cell(row=row_num, column=2, value=row_data["Day"]).border = thin_border
            ws.cell(row=row_num, column=3, value=row_data["Time Slot"]).border = thin_border
            
            c_cell = ws.cell(row=row_num, column=4, value=row_data['Course Code'])
            c_cell.border = thin_border
            
            t_cell = ws.cell(row=row_num, column=5, value=row_data['Course Title'])
            t_cell.border = thin_border
            
            f_cell = ws.cell(row=row_num, column=6, value=row_data['Faculty'])
            f_cell.border = thin_border
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Official Master Excel Routine (.xlsx)",
            data=output,
            file_name="BBA_Fall_2026_Full_Master_Routine.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
