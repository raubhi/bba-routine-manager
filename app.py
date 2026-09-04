import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from ortools.sat.python import cp_model

st.set_page_config(page_title="BBA Routine Optimizer - Grand Matrix", layout="wide")
st.title("BBA Routine Management System | Grand Matrix Engine")
st.divider()

if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}
if 'preview_data' not in st.session_state:
    st.session_state.preview_data = None

# --- 1. DATA UPLOAD & PARSING ---
st.subheader("1. Data Upload & Template Parsing")
course_file = st.file_uploader("Upload Fall 2026 Course Offering Sheet", type=["xlsx"])

if course_file:
    xls = pd.ExcelFile(course_file)
    target_sheet = next((s for s in xls.sheet_names if 'bba' in s.lower() or 'offer' in s.lower()), xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    
    parsed_rows = []
    current_batch = "20th"
    
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
        if not row_vals: continue
        row_str = " ".join(row_vals).lower()
        
        # Capture Batch Headers
        if "batch" in row_str and len(row_vals) < 4:
            for v in row_vals:
                if "batch" in v.lower():
                    current_batch = v.lower().replace("batch", "").strip()
            continue
            
        code_val = next((v for v in row_vals if any(c.isdigit() for c in v) and '-' in v and len(v) >= 5), None)
        
        if code_val and len(row_vals) >= 4:
            code_idx = row_vals.index(code_val)
            title = row_vals[code_idx + 1] if code_idx + 1 < len(row_vals) else "Unknown Course"
            
            sections = ["A"]
            last_val = row_vals[-1]
            if len(last_val) <= 10 and not any(w in last_val.lower() for w in ['semester', 'year', 'cred']):
                sections = [s.strip() for s in last_val.replace(';', ',').split(',') if s.strip()]
                search_space = row_vals[code_idx+2:-1] 
            else:
                search_space = row_vals[code_idx+2:]
                
            teacher = "TBA"
            for cell in reversed(search_space):
                cl = cell.lower()
                if any(w in cl for w in ['semester', 'year', 'cred', 'th', 'st', 'nd', 'rd', 'batch']):
                    continue 
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
    days_ordered = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
    slots_ordered = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]

    for day in days_ordered:
        fac_df[f"{day} Max"] = 3
        
    st.subheader("2. Faculty Roster & Daily Limits Configuration")
    col_config = {"Type": st.column_config.SelectboxColumn("Type", options=["Full-Time", "Adjunct"], required=True)}
    for day in days_ordered:
        col_config[f"{day} Max"] = st.column_config.NumberColumn(f"{day} Max", min_value=0, max_value=6, step=1)
    
    edited_faculty = st.data_editor(fac_df, num_rows="dynamic", column_config=col_config, use_container_width=True)
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
    
    st.divider()

    # --- 2. ADVANCED RULES ---
    st.subheader("3. Advanced Scheduling Constraints")
    col_adv1, col_adv2 = st.columns(2)
    
    with col_adv1:
        st.write("**Specific Faculty Slot Locks**")
        selected_fac = st.selectbox("Select Faculty Member:", available_faculties)
        rule_day = st.selectbox("Select Day:", days_ordered)
        rule_slot = st.selectbox("Select Time Slot:", slots_ordered)
        
        if st.button("➕ Lock Faculty to Slot"):
            st.session_state.fac_rules.append({"Faculty": selected_fac, "Day": rule_day, "Slot": rule_slot})
            st.success(f"Locked {selected_fac} to {rule_day} at {rule_slot}")
            
        if st.session_state.fac_rules:
            st.table(pd.DataFrame(st.session_state.fac_rules))
            if st.button("Clear Faculty Locks"):
                st.session_state.fac_rules = []
                st.rerun()
            
    with col_adv2:
        st.write("**Batch Day Limits & Blackouts**")
        target_batches = st.multiselect("Select Batch(es):", df_tasks['Batch'].unique().tolist())
        batch_max_days = st.number_input(f"Max Active Days:", min_value=1, max_value=6, value=3)
        blocked_batch_days = st.multiselect(f"Strictly Block Days:", days_ordered)
        
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
    
    # --- 4. SOLVER EXECUTION ---
    if st.button("🚀 Generate Full Multi-Batch Master Routine", type="primary", use_container_width=True):
        with st.spinner("Compiling Math Engine and optimizing zero clashes..."):
            
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
            
            x = {}
            for i in range(len(master_tasks)):
                for d in days_ordered:
                    for s in slots_ordered:
                        x[(i, d, s)] = model.NewBoolVar(f"x_{i}_{d}_{s}")
                        
            # BASE CONSTRAINT 1: Assign exactly once
            for i in range(len(master_tasks)):
                model.AddExactlyOne(x[(i, d, s)] for d in days_ordered for s in slots_ordered)
                
            # BASE CONSTRAINT 2 & 3: Clash Prevention (Batch and Faculty)
            for d in days_ordered:
                for s in slots_ordered:
                    batch_sec_map = {}
                    for i, t in enumerate(master_tasks):
                        batch_sec_map.setdefault(f"{t['Batch']} (Sec {t['Section']})", []).append(i)
                    for indices in batch_sec_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices) # GUARANTEES NO COHORT CLASHES
                        
                    fac_map = {}
                    for i, t in enumerate(master_tasks):
                        if t["Faculty"] != "TBA":
                            fac_map.setdefault(t["Faculty"], []).append(i)
                    for indices in fac_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices) # GUARANTEES NO FACULTY CLASHES

            # FACULTY DAILY MAX & OFF-DAYS
            for _, f_row in edited_faculty.iterrows():
                fac_name = f_row["Faculty Name"]
                if fac_name == "TBA": continue
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == fac_name]
                if not f_indices: continue
                for d in days_ordered:
                    max_cap = int(f_row[f"{d} Max"])
                    model.Add(sum(x[(i, d, s)] for i in f_indices for s in slots_ordered) <= max_cap)

            # FACULTY SLOT LOCKS
            for rule in st.session_state.fac_rules:
                r_fac, r_day, r_slot = rule["Faculty"], rule["Day"], rule["Slot"]
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == r_fac]
                if f_indices:
                    model.Add(sum(x[(i, r_day, r_slot)] for i in f_indices) == 1)

            # BATCH DAY LIMITS
            for b_key, rules in st.session_state.batch_rules.items():
                b_indices = [i for i, t in enumerate(master_tasks) if t["Batch"] == b_key]
                if not b_indices: continue
                for d in rules["Blocked Days"]:
                    for i in b_indices:
                        for s in slots_ordered:
                            model.Add(x[(i, d, s)] == 0)
                b_active_vars = []
                for d in days_ordered:
                    day_active = model.NewBoolVar(f"b_active_{b_key}_{d}")
                    model.AddMaxEquality(day_active, [x[(i, d, s)] for i in b_indices for s in slots_ordered])
                    b_active_vars.append(day_active)
                model.Add(sum(b_active_vars) <= rules["Max Days"])

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 15.0
            status = solver.Solve(model)
            
            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                scheduled_output = []
                for i, t in enumerate(master_tasks):
                    for d in days_ordered:
                        for s in slots_ordered:
                            if solver.Value(x[(i, d, s)]) == 1:
                                scheduled_output.append({
                                    "Batch & Section": f"{t['Batch']} (Sec {t['Section']})",
                                    "Day": d,
                                    "Time Slot": s,
                                    "Course Code": t['Code'],
                                    "Course Title": t['Title'],
                                    "Faculty": t['Faculty']
                                })
                st.session_state.preview_data = scheduled_output
                st.success("Routine Generated! Zero Clashes Confirmed.")
            else:
                st.error("❌ Impossible Constraints. Please loosen rules and try again.")

    # --- 5. GRAND MATRIX HTML PREVIEW & EXCEL EXPORT ---
    if st.session_state.preview_data is not None:
        days_ordered = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
        slots_ordered = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]
        
        # Build Grid Dictionary
        batches = sorted(list(set(row["Batch & Section"] for row in st.session_state.preview_data)))
        grid = {b: {d: {s: "" for s in slots_ordered} for d in days_ordered} for b in batches}
        
        for row in st.session_state.preview_data:
            b, d, s = row["Batch & Section"], row["Day"], row["Time Slot"]
            grid[b][d][s] = f"<b>{row['Course Code']}</b><br>{row['Course Title']}<br>{row['Faculty']}"

        st.subheader("📊 Live Grand Matrix Preview")
        st.info("Scroll horizontally to view all days. The solver guarantees zero cohort and faculty clashes in this matrix.")
        
        # Build Custom HTML Table mimicking Google Sheet
        html = "<div style='overflow-x:auto;'><table style='width:100%; border-collapse: collapse; text-align: center; font-size: 13px; font-family: sans-serif;'>"
        
        # Row 1: Days
        html += "<tr><th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>Batch & Section</th>"
        for d in days_ordered:
            html += f"<th colspan='4' style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>{d}</th>"
        html += "</tr>"
        
        # Row 2: Time Slots
        html += "<tr><th style='border: 1px solid #ccc; background-color: #f2f2f2;'></th>"
        for d in days_ordered:
            for s in slots_ordered:
                html += f"<th style='border: 1px solid #ccc; background-color: #e9ecef; padding: 8px;'>{s}</th>"
        html += "</tr>"
        
        # Data Rows
        for b in batches:
            html += f"<tr><td style='border: 1px solid #ccc; font-weight: bold; padding: 10px; background-color: #f8f9fa; white-space: nowrap;'>{b}</td>"
            for d in days_ordered:
                for s in slots_ordered:
                    val = grid[b][d][s]
                    bg = "#ffffff" if val else "#fafafa"
                    html += f"<td style='border: 1px solid #ccc; padding: 10px; background-color: {bg}; min-width: 150px;'>{val}</td>"
            html += "</tr>"
        html += "</table></div>"
        
        # Render HTML directly in Streamlit
        st.markdown(html, unsafe_allow_html=True)
        
        st.divider()
        
        # --- EXCEL EXPORT (EXACT SHEET STRUCTURE) ---
        wb = Workbook()
        ws = wb.active
        ws.title = "Master Routine"
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        
        # Title Row
        ws.merge_cells("A1:Y1")
        title_cell = ws.cell(row=1, column=1, value="UNIVERSITY OF SCHOLARS - BBA PROGRAM ROUTINE (FALL 2026)")
        title_cell.font = Font(name='Arial', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Headers (Batch & Days)
        ws.merge_cells("A2:A3")
        c1 = ws.cell(row=2, column=1, value="Batch & Section")
        c1.fill = header_fill
        c1.font = header_font
        c1.alignment = Alignment(horizontal='center', vertical='center')
        c1.border = thin_border
        
        col_idx = 2
        for day in days_ordered:
            ws.merge_cells(start_row=2, start_column=col_idx, end_row=2, end_column=col_idx+3)
            d_cell = ws.cell(row=2, column=col_idx, value=day)
            d_cell.fill = header_fill
            d_cell.font = header_font
            d_cell.alignment = Alignment(horizontal='center', vertical='center')
            d_cell.border = thin_border
            
            for slot in slots_ordered:
                s_cell = ws.cell(row=3, column=col_idx, value=slot)
                s_cell.fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
                s_cell.font = Font(name='Arial', size=10, bold=True)
                s_cell.alignment = Alignment(horizontal='center', vertical='center')
                s_cell.border = thin_border
                ws.column_dimensions[s_cell.column_letter].width = 20
                col_idx += 1
                
        ws.column_dimensions['A'].width = 25
                
        # Data Rows
        row_idx = 4
        for b in batches:
            ws.cell(row=row_idx, column=1, value=b).border = thin_border
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal='center', vertical='center')
            
            col_idx = 2
            for d in days_ordered:
                for s in slots_ordered:
                    val = grid[b][d][s].replace("<b>", "").replace("</b>", "").replace("<br>", "\n")
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                    cell.border = thin_border
                    col_idx += 1
            row_idx += 1
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Official Master Routine (.xlsx)",
            data=output,
            file_name="BBA_Fall_2026_Grand_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
