import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

def process_uploads(course_file, faculty_file, include_majors):
    # Read the uploaded Excel/CSV files
    if course_file.name.endswith('.csv'):
        df_courses = pd.read_csv(course_file)
    else:
        df_courses = pd.read_excel(course_file)
        
    # Basic sanitization: strip trailing spaces from column names and string data
    df_courses.columns = df_courses.columns.str.strip()
    df_courses = df_courses.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    
    # Filter out major courses if toggle is off
    if not include_majors and 'Major' in df_courses.columns:
        df_courses = df_courses[df_courses['Major'] != 'Yes']
        
    return df_courses.to_dict('records')

def export_routine(solved_data):
    # Generates the exact formatted Excel file for download
    wb = Workbook()
    ws = wb.active
    ws.title = "Optimized Routine"
    
    # Apply strict 3-line cell formatting matching your university template
    for entry in solved_data:
        cell_value = f"{entry['code']}\n{entry['title']}\n{entry['faculty']}"
        cell = ws.cell(row=entry['row'], column=entry['col'])
        cell.value = cell_value
        cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
        cell.font = Font(name='Arial', size=10)
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
