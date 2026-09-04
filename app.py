import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from ortools.sat.python import cp_model

st.set_page_config(page_title="BBA Routine Optimizer - Precise Constraints", layout="wide")
st.title("BBA Routine Management System | Exact Matrix Engine")
st.divider()

# --- INITIALIZE SESSION STATES FOR RULES ---
if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}
if 'preview_df' not in st.session_state:
    st.session_state.preview_df = None

# --- 1. DATA UPLOAD & SMART REVERSE-PARSER ---
st.subheader("1. Data Upload & Template Parsing")
course_file = st.file_uploader("Upload Fall 2026 Course Offering Sheet", type=["xlsx"])

if course_file:
    xls = pd.ExcelFile(course_file)
    target_sheet = next((s for s in xls.sheet_names if 'bba' in s.lower() or 'offer' in s.lower()), xls.sheet_names[0])
    
    # Read raw data to handle merged rows and complex layouts
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    
    parsed_rows = []
    current_batch = "20th"
    
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
        if not row_vals: continue
        row_str = " ".join(row_vals).lower()
        
        # Capture Batch Headers (e.g., "20th Batch")
        if "batch" in row_str and len(row_vals) < 4:
            for v in row_vals:
                if "batch" in v.lower():
                    current_batch = v.lower().replace("batch", "").strip()
            continue
            
        # Identify valid courses via Course Code formatting
        code_val = next((v for v in row_vals if any(c.isdigit() for c in v) and '-' in v and len(v) >= 5), None)
        
        if code_val and len(row_vals) >= 4:
            code_idx = row_vals.index(code_val)
            title = row_vals[code_idx + 1] if code_idx + 1 < len(row_vals) else "Unknown Course"
            
            # Identify Section (usually the last column)
            sections = ["A"]
            last_val = row_vals[-1]
            if len(last_val) <= 10 and not any(w in last_val.lower() for w in ['semester', 'year', 'cred']):
                sections = [s.strip() for s in last_val.replace(';', ',').split(',') if s.strip()]
                search_space = row_vals[code_idx+2:-1] # Search for teacher before section
            else:
                search_space = row_vals[code_idx+2:]
                
            # SMART REVERSE-SEARCH FOR TEACHER (Ignores semesters, years, etc.)
            teacher = "TBA"
            for cell in reversed(search_space):
                cl = cell.lower()
                if any(w in cl for w in ['semester', 'year', 'cred', 'th', 'st', 'nd', 'rd', 'batch']):
                    continue # Skip false positives
                if len(cell) > 2:
                    teacher = cell
                    break
                    
            for sec in sections:
                parsed_rows.append({
                    "Batch": current_batch,
                    "Section": sec,
                    "Code": code_val,
                    "Title": title,
                    "Faculty": teacher
                })

    df_tasks = pd.DataFrame(parsed_rows)

    # Extract Faculty List
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
    days = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
    for day in days:
        fac_df[f"{day} Max"] = 3
        
    st.subheader("2. Faculty Roster & Daily Limits Configuration")
    st.warning("⚠️ **Set a day to 0 to strictly Block it (Off-Day).** Set to 1-6 for maximum allowed classes that day.")
    
    col_config = {"Type": st.column_config.SelectboxColumn("Type", options=["Full-Time", "Adjunct"], required=True)}
    for day in days:
        col_config[f"{day} Max"] = st.column_config.NumberColumn(f"{day} Max", min_value=0, max_value=6, step=1)
    
    edited_faculty = st.data_editor(fac_df, num_rows="dynamic", column_config=col_config, use_container_width=True)
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
    
    st.divider()

    # --- 2. ADVANCED RULES (CONNECTED TO SOLVER) ---
    st.subheader("3. Advanced Scheduling Constraints")
    col_adv1, col_adv2 = st.columns(2)
    
    with col_adv1:
        st.write("**Specific Faculty Slot Locks & Reservations**")
        selected_fac = st.selectbox("Select Faculty Member:", available_faculties)
        rule_day = st.selectbox("Select Day:", days)
        rule_slot = st.selectbox("Select Time Slot:", ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"])
        
        if st.button("➕ Lock Faculty to Slot"):
            st.session_state.fac_rules.append({"Faculty": selected_fac, "Day": rule_day, "Slot": rule_slot})
            st.success(f"Rule Saved: {selected_fac} locked to {rule_day} at {rule_slot}")
            
        if st.session_state.fac_rules:
            st.table(pd.DataFrame(st.session_state.fac_rules))
            if st.button("Clear Faculty Locks"):
                st.session_state.fac_rules = []
                st.rerun()
            
    with col_adv2:
        st.write("**Batch Day Limits & Blackouts**")
        target_batches = st.multiselect("Select Batch(es) to Apply Rules To:", df_tasks['Batch'].unique().tolist())
        batch_max_days = st.number_input(f"Max Active Days for selected batches:", min_value=1, max_value=6, value=3)
        blocked_batch_days = st.multiselect(f"Strictly Block Days for selected batches:", days)
        
        if st.button("💾 Save Batch Configuration"):
            for b in target_batches:
                st.session_state.batch_rules[b] = {"Max Days": batch_max_days, "Blocked Days": blocked_batch_days}
            st.success(f"Configuration saved!")
            
        if st.session_state.batch_rules:
            st.json(st.session_state.batch_rules)
            if st.button("Clear Batch Rules"):
                st.session_state.batch_rules = {}
                st.rerun()

    st.divider()

    # --- 3. GLOBAL ENGINE SETTINGS ---
    st.subheader("4. Global Engine Settings")
    replace_tba = st.toggle("Convert TBA slots to an assigned Faculty member?", value=False)
    tba_replacement = st.selectbox("TBA Replacement Faculty:", available_faculties) if replace_tba else None

    st.divider()
    
    # --- 4. SOLVER EXECUTION (RULES INTEGRATED) ---
    if st.button("🚀 Generate Full Multi-Batch Master Routine", type="primary", use_container_width=True):
        with st.spinner("Compiling UI Constraints and Executing Math Engine..."):
            
            master_tasks = []
            for _, row in df_tasks.iterrows():
                teacher = row['Faculty']
                if teacher.lower() == 'tba' and replace_tba and tba_replacement:
                    teacher = tba_replacement
                    
                master_tasks.append({
                    "Batch": row['Batch'],
                    "Section": row['Section'],
                    "Code": row['Code'],
                    "Title": row['Title'],
                    "Faculty": teacher
                })
            
            model = cp_model.CpModel()
            time_slots = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]
            
            # Initialize Boolean Variables
            x = {}
            for i in range(len(master_tasks)):
                for d in days:
                    for s in time_slots:
                        x[(i, d, s)] = model.NewBoolVar(f"x_{i}_{d}_{s}")
                        
            # BASE CONSTRAINT 1: Assign exactly once
            for i in range(len(master_tasks)):
                model.AddExactlyOne(x[(i, d, s)] for d in days for s in time_slots)
                
            # BASE CONSTRAINT 2 & 3: Clash Prevention (Batch and Faculty)
            for d in days:
                for s in time_slots:
                    # Cohort Clash
                    batch_sec_map = {}
                    for i, t in enumerate(master_tasks):
                        batch_sec_map.setdefault((t["Batch"], t["Section"]), []).append(i)
                    for indices in batch_sec_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices)
                        
                    # Faculty Clash
                    fac_map = {}
                    for i, t in enumerate(master_tasks):
                        if t["Faculty"] != "TBA":
                            fac_map.setdefault(t["Faculty"], []).append(i)
                    for indices in fac_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices)

            # --- INTEGRATING UI RULE 1: FACULTY DAILY MAX & OFF-DAYS ---
            for _, f_row in edited_faculty.iterrows():
                fac_name = f_row["Faculty Name"]
                if fac_name == "TBA": continue
                
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == fac_name]
                if not f_indices: continue
                
                for d in days:
                    max_cap = int(f_row[f"{d} Max"])
                    # If max_cap is 0, this strictly acts as an OFF-DAY block.
                    model.Add(sum(x[(i, d, s)] for i in f_indices for s in time_slots) <= max_cap)

            # --- INTEGRATING UI RULE 2: FACULTY SLOT LOCKS ---
            for rule in st.session_state.fac_rules:
                r_fac, r_day, r_slot = rule["Faculty"], rule["Day"], rule["Slot"]
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == r_fac]
                if f_indices:
                    # Forces the solver to place exactly one of this faculty's classes into the locked slot
                    model.Add(sum(x[(i, r_day, r_slot)] for i in f_indices) == 1)

            # --- INTEGRATING UI RULE 3: BATCH DAY LIMITS & BLACKOUTS ---
            for b_key, rules in st.session_state.batch_rules.items():
                b_indices = [i for i, t in enumerate(master_tasks) if t["Batch"] == b_key]
                if not b_indices: continue
                
                # Apply Blackouts
                for d in rules["Blocked Days"]:
                    for i in b_indices:
                        for s in time_slots:
                            model.Add(x[(i, d, s)] == 0)
                            
                # Apply Active Day Limits (e.g., Max 3 days a week)
                b_active_vars = []
                for d in days:
                    day_active = model.NewBoolVar(f"b_active_{b_key}_{d}")
                    model.AddMaxEquality(day_active, [x[(i, d, s)] for i in b_indices for s in time_slots])
                    b_active_vars.append(day_active)
                model.Add(sum(b_active_vars) <= rules["Max Days"])

            # Execute Solver
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 15.0
            status = solver.Solve(model)
            
            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                scheduled_output = []
                for i, t in enumerate(master_tasks):
                    for d in days:
                        for s in time_slots:
                            if solver.Value(x[(i, d, s)]) == 1:
                                scheduled_output.append({
                                    "Batch & Section": f"{t['Batch']} (Sec {t['Section']})",
                                    "Day": d,
                                    "Time Slot": s,
                                    "Course Code": t['Code'],
                                    "Course Title": t['Title'],
                                    "Faculty": t['Faculty']
                                })
                
                # Order by Day and Time Slot for visual clarity
                df_out = pd.DataFrame(scheduled_output)
                day_order = {"Saturday":1, "Sunday":2, "Monday":3, "Tuesday":4, "Wednesday":5, "Thursday":6}
                df_out['Day_Num'] = df_out['Day'].map(day_order)
                df_out = df_out.sort_values(by=['Batch & Section', 'Day_Num', 'Time Slot']).drop('Day_Num', axis=1)
                
                st.session_state.preview_df = df_out
                st.success("All reservations and batch limits strictly applied. Routine Generated!")
            else:
                st.error("❌ Schedule is mathematically impossible. You have locked too many constraints (e.g., locking a teacher to an off-day, or capping batch days too tightly). Please clear some rules and try again.")

    # --- 5. VISUAL GRID VIEW & EXCEL EXPORT (EXACT FORMAT MATCH) ---
    if st.session_state.preview_df is not None:
        st.subheader("📊 Routine Preview Window")
        
        view_mode = st.radio("Select View:", ["Master List View (Matches Exact Format)", "Visual Grid Matrix"], horizontal=True)
        
        if view_mode == "Master List View (Matches Exact Format)":
            st.dataframe(st.session_state.preview_df, use_container_width=True, hide_index=True)
        else:
            df_p = st.session_state.preview_df.copy()
            df_p['Cell'] = df_p['Course Code'] + "\n" + df_p['Faculty']
            batch_sel = st.selectbox("Inspect Batch on Grid:", df_p['Batch & Section'].unique())
            pivot = df_p[df_p['Batch & Section'] == batch_sel].pivot_table(index="Time Slot", columns="Day", values="Cell", aggfunc=lambda x: ' | '.join(x))
            st.dataframe(pivot, use_container_width=True)
        
        # Build Exact Format Master OpenPyXL Excel File
        wb = Workbook()
        ws = wb.active
        ws.title = "Master Routine"
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="000000")
        
        # Exact headers from your image reference
        headers = ["Batch & Section", "Day", "Time Slot", "Course Code", "Course Title", "Faculty"]
        for col_num, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = thin_border
            ws.column_dimensions[cell.column_letter].width = 25
            
        for r_idx, row_data in st.session_state.preview_df.iterrows():
            row_num = r_idx + 2
            for c_idx, col_name in enumerate(headers, 1):
                cell = ws.cell(row=row_num, column=c_idx, value=row_data[col_name])
                cell.border = thin_border
                cell.font = Font(name='Arial', size=10)
                cell.alignment = Alignment(horizontal='left', vertical='center')
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Official Routine (Matches Image Format)",
            data=output,
            file_name="BBA_Fall_2026_Routine.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
