import streamlit as st
import pandas as pd
import json
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from ortools.sat.python import cp_model
import math

st.set_page_config(page_title="BBA Routine Optimizer - Complete Matrix", layout="wide")
st.title("BBA Routine Management System | Complete Matrix Engine")
st.divider()

# --- INITIALIZE PERSISTENT SESSION STATES ---
if 'fac_rules' not in st.session_state:
    st.session_state.fac_rules = []
if 'batch_rules' not in st.session_state:
    st.session_state.batch_rules = {}
if 'sun_mon_facs' not in st.session_state:
    st.session_state.sun_mon_facs = []
if 'preview_data' not in st.session_state:
    st.session_state.preview_data = None
if 'current_file_id' not in st.session_state:
    st.session_state.current_file_id = None
if 'base_fac_df' not in st.session_state:
    st.session_state.base_fac_df = None

# --- CONFIGURATION MANAGER (SAVE/LOAD) ---
st.subheader("💾 Configuration Manager (Save/Load Rules)")
st.info("Save your constraints locally so you never have to re-enter them if the app restarts.")
col_conf1, col_conf2 = st.columns(2)
with col_conf1:
    config_dict = {
        "fac_rules": st.session_state.fac_rules,
        "batch_rules": st.session_state.batch_rules,
        "sun_mon_facs": st.session_state.sun_mon_facs
    }
    st.download_button(
        label="📥 Download Saved Rules (.json)",
        data=json.dumps(config_dict),
        file_name="bba_routine_rules.json",
        mime="application/json"
    )
with col_conf2:
    uploaded_config = st.file_uploader("📤 Upload Saved Rules (.json)", type=["json"])
    if uploaded_config:
        loaded_conf = json.load(uploaded_config)
        st.session_state.fac_rules = loaded_conf.get("fac_rules", [])
        st.session_state.batch_rules = loaded_conf.get("batch_rules", {})
        st.session_state.sun_mon_facs = loaded_conf.get("sun_mon_facs", [])
        st.success("Rules successfully loaded! They will apply to the next generation.")

st.divider()

# --- 1. DATA UPLOAD & PARSING ---
st.subheader("1. Data Upload & Template Parsing")
course_file = st.file_uploader("Upload Fall 2026 Course Offering Sheet", type=["xlsx"])

if course_file:
    file_id = f"{course_file.name}_{course_file.size}"
    is_new_file = False
    if st.session_state.current_file_id != file_id:
        st.session_state.current_file_id = file_id
        is_new_file = True

    xls = pd.ExcelFile(course_file)
    target_sheet = next((s for s in xls.sheet_names if 'bba' in s.lower() or 'offer' in s.lower()), xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    
    parsed_rows = []
    current_batch = "20th"
    
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
        if not row_vals: continue
        row_str = " ".join(row_vals).lower()
        
        if "batch" in row_str and len(row_vals) < 4:
            for v in row_vals:
                if "batch" in v.lower():
                    current_batch = v.lower().replace("batch", "").strip()
            continue
            
        code_val = next((v for v in row_vals if any(c.isdigit() for c in v) and '-' in v and len(v) >= 5), None)
        
        if code_val and len(row_vals) >= 4:
            code_idx = row_vals.index(code_val)
            title = row_vals[code_idx + 1] if code_idx + 1 < len(row_vals) else "Unknown Course"
            
            # FIXED UNLIMITED SECTION PARSER
            sections = ["A"]
            last_val = row_vals[-1]
            if not any(w in last_val.lower() for w in ['semester', 'year', 'cred', 'total', 'th', 'st', 'nd', 'rd', 'batch']):
                sec_clean = last_val.replace(';', ',').replace('&', ',').replace('and', ',')
                sections = [s.strip() for s in sec_clean.split(',') if s.strip()]
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
                    "Batch": f"{current_batch} (Sec {sec})",
                    "Code": code_val,
                    "Title": title,
                    "Faculty": teacher
                })

    df_tasks = pd.DataFrame(parsed_rows)

    if is_new_file or st.session_state.base_fac_df is None:
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
        for day in ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]:
            fac_df[f"{day} Max"] = 3
            
        st.session_state.base_fac_df = fac_df

    days_ordered = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
    slots_ordered = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]

    st.subheader("2. Persistent Faculty Roster & Daily Limits Matrix")
    col_config = {"Type": st.column_config.SelectboxColumn("Type", options=["Full-Time", "Adjunct"], required=True)}
    for day in days_ordered:
        col_config[f"{day} Max"] = st.column_config.NumberColumn(f"{day} Max", min_value=0, max_value=6, step=1)
    
    edited_faculty = st.data_editor(st.session_state.base_fac_df, num_rows="dynamic", column_config=col_config, use_container_width=True)
    st.session_state.base_fac_df = edited_faculty 
    
    available_faculties = edited_faculty[edited_faculty['Faculty Name'] != 'TBA']['Faculty Name'].tolist()
    
    st.divider()

    # --- 2. ADVANCED RULES (TABBED ENVIRONMENT) ---
    st.subheader("3. Advanced Scheduling Constraints")
    tab_locks, tab_batches, tab_sunmon = st.tabs(["Specific Slot Locks", "Batch Limits & Blackouts", "Sun/Mon Forced Load"])
    
    with tab_locks:
        st.write("**Specific Faculty Slot Locks**")
        col_t1, col_t2 = st.columns([1, 2])
        with col_t1:
            selected_fac = st.selectbox("Select Faculty:", available_faculties, key="l_fac")
            rule_day = st.selectbox("Select Day:", days_ordered, key="l_day")
            rule_slot = st.selectbox("Select Time Slot:", slots_ordered, key="l_slot")
            if st.button("➕ Lock Faculty to Slot"):
                st.session_state.fac_rules.append({"Faculty": selected_fac, "Day": rule_day, "Slot": rule_slot})
                st.rerun()
        with col_t2:
            if st.session_state.fac_rules:
                st.table(pd.DataFrame(st.session_state.fac_rules))
                if st.button("Clear All Faculty Locks"):
                    st.session_state.fac_rules = []
                    st.rerun()
            
    with tab_batches:
        st.write("**Batch Day Limits & Blackouts**")
        col_b1, col_b2 = st.columns([1, 2])
        with col_b1:
            target_batches = st.multiselect("Select Batch(es):", df_tasks['Batch'].unique().tolist())
            batch_max_days = st.number_input(f"Max Active Days:", min_value=1, max_value=6, value=3)
            blocked_batch_days = st.multiselect(f"Strictly Block Days:", days_ordered)
            if st.button("💾 Save Batch Configuration"):
                for b in target_batches:
                    st.session_state.batch_rules[b] = {"Max Days": batch_max_days, "Blocked Days": blocked_batch_days}
                st.rerun()
        with col_b2:
            if st.session_state.batch_rules:
                st.json(st.session_state.batch_rules)
                if st.button("Clear All Batch Rules"):
                    st.session_state.batch_rules = {}
                    st.rerun()

    with tab_sunmon:
        st.write("**Force 3 Classes on Sunday & Monday**")
        col_sm1, col_sm2 = st.columns([1, 2])
        with col_sm1:
            sm_fac = st.selectbox("Select Faculty:", available_faculties, key="sm_fac_sel")
            if st.button("➕ Force Sun/Mon 3-Class Load"):
                if sm_fac not in st.session_state.sun_mon_facs:
                    st.session_state.sun_mon_facs.append(sm_fac)
                    st.rerun()
        with col_sm2:
            if st.session_state.sun_mon_facs:
                st.write("**Currently Forced Faculties:**")
                st.write(st.session_state.sun_mon_facs)
                if st.button("Clear Sun/Mon Locks"):
                    st.session_state.sun_mon_facs = []
                    st.rerun()

    st.divider()

    # --- 3. GLOBAL ENGINE SETTINGS ---
    st.subheader("4. Global Engine Settings")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        include_majors = st.toggle("Include Major Courses (Uncheck to drop)", value=False)
        compact_schedule = st.toggle("Enforce Compact Scheduling (Prevent extreme gaps)", value=True)
        replace_tba = st.toggle("Convert TBA slots to an assigned Faculty member?", value=False)
        tba_replacement = st.selectbox("TBA Replacement Faculty:", available_faculties) if replace_tba else None
    with col_g2:
        total_rooms = st.number_input("Total Available Rooms per Time Slot:", min_value=1, max_value=30, value=6, step=1)

    st.divider()
    
    # --- 4. PRE-DIAGNOSTICS & SOLVER EXECUTION ---
    if st.button("🚀 Generate Full Multi-Batch Master Routine", type="primary", use_container_width=True):
        
        master_tasks = []
        for _, row in df_tasks.iterrows():
            if not include_majors and "major" in row['Title'].lower():
                continue
            teacher = row['Faculty']
            if teacher.lower() == 'tba' and replace_tba and tba_replacement:
                teacher = tba_replacement
            master_tasks.append({
                "Batch": row['Batch'],
                "Code": row['Code'],
                "Title": row['Title'],
                "Faculty": teacher
            })
            
        # --- VISUAL PRE-GENERATION DIAGNOSTICS ---
        fac_counts = {}
        for t in master_tasks:
            fac_counts[t["Faculty"]] = fac_counts.get(t["Faculty"], 0) + 1
            
        diagnostic_errors = []
        
        # 1. Check Sun/Mon rule viability
        for sm_fac in st.session_state.sun_mon_facs:
            assigned = fac_counts.get(sm_fac, 0)
            if assigned < 6:
                diagnostic_errors.append(f"👨‍🏫 **{sm_fac}** is locked to take 6 classes on Sun/Mon, but the Excel sheet only assigns them **{assigned}** section(s).")
                
        # 2. Check total room capacity
        total_slots = len(days_ordered) * len(slots_ordered)
        max_possible_classes = total_slots * total_rooms
        if len(master_tasks) > max_possible_classes:
            diagnostic_errors.append(f"🏢 **Not Enough Rooms:** You have **{len(master_tasks)}** total sections to schedule, but {total_rooms} rooms only allows a maximum of **{max_possible_classes}** classes per week.")
            
        if diagnostic_errors:
            st.error("❌ **Pre-Generation Diagnostic Failed:** Please fix these mathematical contradictions before running the solver.")
            for e in diagnostic_errors:
                st.warning(e)
            st.stop()
            
        # --- MATH SOLVER ---
        with st.spinner("Diagnostics Passed. Compiling Math Engine and Optimizing..."):
            model = cp_model.CpModel()
            x = {}
            for i in range(len(master_tasks)):
                for d in days_ordered:
                    for s in slots_ordered:
                        x[(i, d, s)] = model.NewBoolVar(f"x_{i}_{d}_{s}")
                        
            for i in range(len(master_tasks)):
                model.AddExactlyOne(x[(i, d, s)] for d in days_ordered for s in slots_ordered)
                
            for d in days_ordered:
                for s in slots_ordered:
                    model.Add(sum(x[(i, d, s)] for i in range(len(master_tasks))) <= total_rooms)
                    
                    batch_sec_map = {}
                    for i, t in enumerate(master_tasks):
                        batch_sec_map.setdefault(t['Batch'], []).append(i)
                    for indices in batch_sec_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices)
                        
                    fac_map = {}
                    for i, t in enumerate(master_tasks):
                        if t["Faculty"] != "TBA":
                            fac_map.setdefault(t["Faculty"], []).append(i)
                    for indices in fac_map.values():
                        model.AddAtMostOne(x[(i, d, s)] for i in indices)

            active_facs = list(set(t["Faculty"] for t in master_tasks if t["Faculty"] != "TBA"))
            for _, f_row in edited_faculty.iterrows():
                fac_name = f_row["Faculty Name"]
                if fac_name == "TBA": continue
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == fac_name]
                if not f_indices: continue
                for d in days_ordered:
                    max_cap = int(f_row[f"{d} Max"])
                    model.Add(sum(x[(i, d, s)] for i in f_indices for s in slots_ordered) <= max_cap)

            for rule in st.session_state.fac_rules:
                r_fac, r_day, r_slot = rule["Faculty"], rule["Day"], rule["Slot"]
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == r_fac]
                if f_indices:
                    model.Add(sum(x[(i, r_day, r_slot)] for i in f_indices) == 1)

            for sm_fac in st.session_state.sun_mon_facs:
                f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == sm_fac]
                if f_indices:
                    model.Add(sum(x[(i, "Sunday", s)] for i in f_indices for s in slots_ordered) == 3)
                    model.Add(sum(x[(i, "Monday", s)] for i in f_indices for s in slots_ordered) == 3)

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

            if compact_schedule:
                for f in active_facs:
                    f_indices = [i for i, t in enumerate(master_tasks) if t["Faculty"] == f]
                    for d in days_ordered:
                        s1 = sum(x[(i, d, slots_ordered[0])] for i in f_indices)
                        s2 = sum(x[(i, d, slots_ordered[1])] for i in f_indices)
                        s3 = sum(x[(i, d, slots_ordered[2])] for i in f_indices)
                        s4 = sum(x[(i, d, slots_ordered[3])] for i in f_indices)
                        model.Add(s1 + s4 <= 1 + s2 + s3)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 20.0
            status = solver.Solve(model)
            
            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                scheduled_output = []
                for i, t in enumerate(master_tasks):
                    for d in days_ordered:
                        for s in slots_ordered:
                            if solver.Value(x[(i, d, s)]) == 1:
                                scheduled_output.append({
                                    "Batch": t['Batch'],
                                    "Day": d,
                                    "Time Slot": s,
                                    "Course Code": t['Code'],
                                    "Course Title": t['Title'],
                                    "Faculty": t['Faculty']
                                })
                st.session_state.preview_data = scheduled_output
                st.success(f"Routine Generated! Scheduled {len(master_tasks)} courses across {total_rooms} rooms.")
            else:
                st.error("❌ Impossible Constraints. Ensure batch blackouts aren't forcing too many classes into a single day.")

    # --- 5. TABBED VIEWS & EXCEL EXPORT ---
    if st.session_state.preview_data is not None:
        days_ordered = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
        slots_ordered = ["9:30-11:00", "11:00-12:30", "1:30-3:00", "3:00-4:30"]
        batches = sorted(list(set(row["Batch"] for row in st.session_state.preview_data)))
        
        tab1, tab2 = st.tabs(["📊 Master Grid View", "👨‍🏫 Faculty Individual Routines"])
        
        with tab1:
            st.info("Days and Time Slots on the left, Batch Names across the top.")
            
            grid = {d: {s: {b: "" for b in batches} for s in slots_ordered} for d in days_ordered}
            for row in st.session_state.preview_data:
                b, d, s = row["Batch"], row["Day"], row["Time Slot"]
                grid[d][s][b] = f"<b>{row['Course Code']}</b><br>{row['Course Title']}<br>{row['Faculty']}"

            html = "<div style='overflow-x:auto;'><table style='width:100%; border-collapse: collapse; text-align: center; font-size: 13px; font-family: sans-serif;'>"
            html += "<tr><th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>Day</th>"
            html += "<th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>Time Slot</th>"
            for b in batches:
                html += f"<th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>{b}</th>"
            html += "</tr>"
            
            for d in days_ordered:
                for i, s in enumerate(slots_ordered):
                    html += "<tr>"
                    if i == 0:
                        html += f"<td rowspan='4' style='border: 1px solid #ccc; font-weight: bold; background-color: #f8f9fa; vertical-align: middle;'>{d}</td>"
                    html += f"<td style='border: 1px solid #ccc; background-color: #e9ecef; padding: 8px; white-space: nowrap;'>{s}</td>"
                    for b in batches:
                        val = grid[d][s][b]
                        bg = "#ffffff" if val else "#fafafa"
                        html += f"<td style='border: 1px solid #ccc; padding: 10px; background-color: {bg}; min-width: 150px;'>{val}</td>"
                    html += "</tr>"
            html += "</table></div>"
            st.markdown(html, unsafe_allow_html=True)
            
            st.divider()
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Master Routine"
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
            
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(batches)+2)
            title_cell = ws.cell(row=1, column=1, value="UNIVERSITY OF SCHOLARS - BBA PROGRAM ROUTINE (FALL 2026)")
            title_cell.font = Font(name='Arial', size=14, bold=True)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            ws.cell(row=2, column=1, value="Day").fill = header_fill
            ws.cell(row=2, column=1).font = header_font
            ws.cell(row=2, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=2, column=1).border = thin_border
            
            ws.cell(row=2, column=2, value="Time Slot").fill = header_fill
            ws.cell(row=2, column=2).font = header_font
            ws.cell(row=2, column=2).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=2, column=2).border = thin_border
            
            for idx, b in enumerate(batches, 3):
                c = ws.cell(row=2, column=idx, value=b)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = thin_border
                ws.column_dimensions[c.column_letter].width = 25
                
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15
            
            row_idx = 3
            for d in days_ordered:
                start_row = row_idx
                for s in slots_ordered:
                    sc = ws.cell(row=row_idx, column=2, value=s)
                    sc.alignment = Alignment(horizontal='center', vertical='center')
                    sc.border = thin_border
                    
                    for c_idx, b in enumerate(batches, 3):
                        val = grid[d][s][b].replace("<b>", "").replace("</b>", "").replace("<br>", "\n")
                        cell = ws.cell(row=row_idx, column=c_idx, value=val)
                        cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                        cell.border = thin_border
                    row_idx += 1
                    
                ws.merge_cells(start_row=start_row, start_column=1, end_row=row_idx-1, end_column=1)
                dc = ws.cell(row=start_row, column=1, value=d)
                dc.alignment = Alignment(horizontal='center', vertical='center')
                dc.font = Font(bold=True)
                for r in range(start_row, row_idx):
                    ws.cell(row=r, column=1).border = thin_border
                
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 Download Official Master Routine (.xlsx)",
                data=output,
                file_name="BBA_Fall_2026_Master_Routine.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        with tab2:
            st.info("Select one or more faculty members to view their specific schedules side-by-side in a single grid.")
            active_facs = sorted(list(set(row["Faculty"] for row in st.session_state.preview_data)))
            selected_facs = st.multiselect("Select Faculty to compare:", active_facs)
            
            if selected_facs:
                fac_grid = {d: {s: {f: "" for f in selected_facs} for s in slots_ordered} for d in days_ordered}
                for row in st.session_state.preview_data:
                    fac = row["Faculty"]
                    if fac in selected_facs:
                        fac_grid[row["Day"]][row["Time Slot"]][fac] = f"<b>{row['Course Code']}</b><br>{row['Batch']}"

                fac_html = "<div style='overflow-x:auto;'><table style='width:100%; border-collapse: collapse; text-align: center; font-size: 13px; font-family: sans-serif;'>"
                fac_html += "<tr><th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>Day</th>"
                fac_html += "<th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>Time Slot</th>"
                for f in selected_facs:
                    fac_html += f"<th style='border: 1px solid #ccc; background-color: #1F4E78; color: white; padding: 10px;'>{f}</th>"
                fac_html += "</tr>"
                
                for d in days_ordered:
                    for i, s in enumerate(slots_ordered):
                        fac_html += "<tr>"
                        if i == 0:
                            fac_html += f"<td rowspan='4' style='border: 1px solid #ccc; font-weight: bold; background-color: #f8f9fa; vertical-align: middle;'>{d}</td>"
                        fac_html += f"<td style='border: 1px solid #ccc; background-color: #e9ecef; padding: 8px; white-space: nowrap;'>{s}</td>"
                        for f in selected_facs:
                            val = fac_grid[d][s][f]
                            bg = "#e6f2ff" if val else "#ffffff"
                            fac_html += f"<td style='border: 1px solid #ccc; padding: 10px; background-color: {bg}; min-width: 150px;'>{val}</td>"
                        fac_html += "</tr>"
                fac_html += "</table></div>"
                st.markdown(fac_html, unsafe_allow_html=True)
