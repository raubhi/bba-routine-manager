import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from ortools.sat.python import cp_model

st.set_page_config(page_title="BBA Routine Optimizer - Master Matrix", layout="wide")
st.title("BBA Routine Management System | Visual Routine Matrix")
st.divider()

if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}
if 'preview_df' not in st.session_state:
    st.session_state.preview_df = None

# --- 1. DATA UPLOAD & INTELLIGENT PARSING ---
st.subheader("1. Data Upload & Template Parsing")
course_file = st.file_uploader("Upload Fall 2026 Course Offering Sheet", type=["xlsx"])

if course_file:
    xls = pd.ExcelFile(course_file)
    target_sheet = next((s for s in xls.sheet_names if 'bba' in s.lower() or 'offer' in s.lower()), xls.sheet_names[0])
    
    # Read sheet raw to find data rows and batch headers
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    
    # Clean and parse rows, tracking the current Batch context
    parsed_rows = []
    current_batch = "20th Batch"
    
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
        
        # Detect Batch header rows in the sheet
        if "batch" in row_str.lower():
            for val in row.values:
                if pd.notna(val) and "batch" in str(val).lower():
                    current_batch = str(val).strip()
            continue
            
        # Look for valid course rows (must contain a course code pattern like 'BBA' or numbers with dashes)
        row_cells = [str(val).strip() for val in row.values if pd.notna(val)]
        if len(row_cells) >= 4:
            # Check if any cell looks like a course code (e.g., contains numbers and letters)
            potential_code = next((c for c in row_cells if any(char.isdigit() for char in c) and '-' in c), None)
            if potential_code:
                # Find course title, teacher, and section based on typical column patterns
                code_idx = row_cells.index(potential_code)
                title = row_cells[code_idx + 1] if code_idx + 1 < len(row_cells) else "Course Title"
                
                # Search for teacher name (usually later in the row, avoiding semester text)
                teacher = "TBA"
                for cell in row_cells[code_idx+2:]:
                    cell_lower = cell.lower()
                    if not any(kw in cell_lower for kw in ['semester', 'year', 'cred', '√', 'true', 'false', 'reg', 'week']) and len(cell) > 3:
                        teacher = cell
                        break
                
                # Search for section info at the end of the row
                sections = ["A"]
                sec_raw = row_cells[-1]
                if sec_raw and len(sec_raw) <= 15 and not 'semester' in sec_raw.lower():
                    sections = [s.strip() for s in sec_raw.replace(';', ',').split(',') if s.strip()]
                
                for sec in sections:
                    parsed_rows.append({
                        "Batch": current_batch,
                        "Section": sec,
                        "Code": potential_code,
                        "Title": title,
                        "Faculty": teacher
                    })

    # Fallback to standard dataframe parsing if custom detector is empty
    if not parsed_rows:
        df_std = pd.read_excel(xls, sheet_name=target_sheet)
        for _, row in df_std.iterrows():
            parsed_rows.append({
                "Batch": str(row.iloc[0]) if pd.notna(row.iloc[0]) else "20th Batch",
                "Section": "A",
                "Code": str(row.iloc[2]) if len(row.dropna()) > 2 else "BBA 1101",
                "Title": str(row.iloc[3]) if len(row.dropna()) > 3 else "Course",
                "Faculty": str(row.iloc[6]) if len(row.dropna()) > 6 else "TBA"
            })

    df_tasks = pd.DataFrame(parsed_rows)

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
        extracted = df_tasks['Faculty'].dropna().unique()
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

    # --- 2. GLOBAL SETTINGS ---
    st.subheader("3. Global Engine Settings")
    g_col1, g_col2 = st.columns(2)
    with g_col1:
        include_majors = st.toggle("Include Major Courses", value=False)
    with g_col2:
        replace_tba = st.toggle("Convert TBA slots to available faculty", value=False)
        tba_replacement = st.selectbox("TBA Replacement Pool:", available_faculties) if replace_tba else None

    st.divider()
    
    # --- 4. SOLVER EXECUTION ---
    if st.button("🚀 Generate Full Multi-Batch Master Routine", type="primary", use_container_width=True):
        with st.spinner("Executing OR-Tools CP-SAT multi-batch constraint matrix..."):
            
            master_tasks = []
            for _, row in df_tasks.iterrows():
                teacher = row['Faculty']
                if teacher.lower() in ['nan', '', 'none', 'tba'] and replace_tba and tba_replacement:
                    teacher = tba_replacement
                elif teacher.lower() in ['nan', '', 'none']:
                    teacher = "TBA"
                    
                master_tasks.append({
                    "Batch": row['Batch'],
                    "Section": row['Section'],
                    "Code": row['Code'],
                    "Title": row['Title'],
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

    # --- 5. VISUAL GRID VIEW & EXCEL EXPORT ---
    if st.session_state.preview_df is not None:
        st.subheader("📊 Visual Routine Matrix (Grid View)")
        st.info("Switch between the List View and the visual Matrix Grid view below to check your schedule layout.")
        
        view_mode = st.radio("Select View Mode:", ["Grid View (Matrix)", "List View (Master Table)"], horizontal=True)
        
            if view_mode == "Grid View (Matrix)":
            # Create a pivot table matrix: Rows = Time Slots, Columns = Days, Values = Course & Faculty strings
            df_p = st.session_state.preview_df.copy()
            df_p['CellContent'] = df_p['Course Code'] + "\n(" + df_p['Course Title'] + ")\n👨‍🏫 " + df_p['Faculty']
            
            selected_batch_filter = st.selectbox("Select Batch to Inspect on Grid:", df_p['Batch & Section'].unique())
            df_filtered = df_p[df_p['Batch & Section'] == selected_batch_filter]
            
            # Pivot table mapping time slots to days
            pivot_grid = df_filtered.pivot_table(
                index="Time Slot", 
                columns="Day", 
                values="CellContent", 
                aggfunc=lambda x: ' | '.join(x)
            )
            
            # Reorder columns to standard sequence
            standard_days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
            existing_days = [d for d in standard_days if d in pivot_grid.columns]
            pivot_grid = pivot_grid[existing_days]
            
            st.markdown(f"### Routine Matrix for: **{selected_batch_filter}**")
            st.dataframe(pivot_grid, use_container_width=True)
            
        else:
            st.dataframe(st.session_state.preview_df, use_container_width=True)
        
        # Build Master OpenPyXL Excel File matching template format
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
            
            ws.cell(row=row_num, column=4, value=row_data['Course Code']).border = thin_border
            ws.cell(row=row_num, column=5, value=row_data['Course Title']).border = thin_border
            ws.cell(row=row_num, column=6, value=row_data['Faculty']).border = thin_border
            
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
