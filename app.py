import streamlit as st
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import Rule, IconSet, FormatObject, CellIsRule
import re
import io
import zipfile

def get_class_info_from_sheet(sheet):
    students = []
    advisor_name = ""
    start_reading = False
    
    for row in sheet.iter_rows(values_only=True):
        if row[1] == "STUDENT NUMBER":
            start_reading = True
            for cell_val in row:
                if cell_val and isinstance(cell_val, str) and "Advisor" in cell_val:
                    advisor_name = cell_val.split(":")[-1].strip()
                    break
            continue
            
        if start_reading:
            if not row[0] or not str(row[0]).strip().isdigit():
                break
            students.append({
                "index": row[0],
                "number": row[1],
                "name": row[2],
                "surname": row[3]
            })
            
    return students, advisor_name

def get_template_student_rows(wb, sheet_idx, start_row=3):
    ws = wb.worksheets[sheet_idx]
    
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row <= start_row <= max_row:
            return max_row - start_row + 1
            
    count = 0
    for r in range(start_row, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip() != "" and str(val).strip() != "0":
            count += 1
        else:
            break
            
    if count > 0:
        return count
        
    if sheet_idx > 0:
        return get_template_student_rows(wb, 0, start_row)
        
    return 30

def shift_formula_rows(formula_str, threshold_row, offset):
    if not formula_str or not isinstance(formula_str, str) or not formula_str.startswith('='):
        return formula_str
        
    def repl(match):
        prefix = match.group(1)
        row_num = int(match.group(2))
        if row_num >= threshold_row:
            return f"{prefix}{row_num + offset}"
        return match.group(0)
        
    return re.sub(r'(?<![A-Za-z])(\$?[A-Za-z]{1,3}\$?)(\d+)\b', repl, formula_str)

def adjust_template_rows_and_tables(ws, num_students, current_rows):
    start_row = 3
    original_last_student_row = start_row + current_rows - 1
    
    action_row_idx = start_row + (current_rows // 2)
    if action_row_idx <= start_row:
        action_row_idx = start_row + 1
    
    offset = 0
    if num_students > current_rows:
        rows_to_add = num_students - current_rows
        ws.insert_rows(action_row_idx, amount=rows_to_add)
        offset = rows_to_add
        
        for r in range(action_row_idx, action_row_idx + rows_to_add):
            for col in range(1, ws.max_column + 1):
                source_cell = ws.cell(row=start_row, column=col)
                target_cell = ws.cell(row=r, column=col)
                if source_cell.has_style:
                    target_cell._style = source_cell._style

    elif num_students < current_rows:
        rows_to_delete = current_rows - num_students
        ws.delete_rows(action_row_idx, amount=rows_to_delete)
        offset = -rows_to_delete

    last_student_row = start_row + num_students - 1

    if offset != 0:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    cell.value = shift_formula_rows(str(cell.value), action_row_idx, offset)

    for table in ws.tables.values():
        ref = table.ref
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        table_offset = max_row - original_last_student_row
        if table_offset < 0:
            table_offset = 0
        new_table_max_row = last_student_row + table_offset
        table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_table_max_row}"

    for col in range(5, ws.max_column + 1):
        master_cell = ws.cell(row=start_row, column=col)
        if master_cell.data_type != 'f':
            master_cell.value = None

    for r in range(start_row + 1, last_student_row + 1):
        for col in range(1, ws.max_column + 1):
            master_cell = ws.cell(row=start_row, column=col)
            target_cell = ws.cell(row=r, column=col)
            
            if master_cell.data_type == 'f' and master_cell.value:
                try:
                    target_cell.value = Translator(master_cell.value, origin=master_cell.coordinate).translate_formula(target_cell.coordinate)
                except:
                    target_cell.value = master_cell.value
            else:
                if col >= 5:
                    target_cell.value = None

    if hasattr(ws, 'conditional_formatting') and hasattr(ws.conditional_formatting, '_cf_rules'):
        ws.conditional_formatting._cf_rules.clear()
        
    if hasattr(ws, 'extLst'):
        ws.extLst = None

    return last_student_row 

def process_class_template(template_bytes, class_name, students, module_name, advisor_name):
    wb = openpyxl.load_workbook(filename=io.BytesIO(template_bytes))
    wb.template = False 
    
    first_sheet_last_row = 3
    
    for i, sheet_name in enumerate(wb.sheetnames):
        ws = wb[sheet_name]
        template_student_rows = get_template_student_rows(wb, i, start_row=3)
        last_student_row = adjust_template_rows_and_tables(ws, len(students), template_student_rows)
        
        if i == 0:
            first_sheet_last_row = last_student_row
            
        if i > 0:
            cfvo1 = FormatObject(type='num', val=0)   
            cfvo2 = FormatObject(type='num', val=45)  
            cfvo3 = FormatObject(type='num', val=60)  
            cfvo4 = FormatObject(type='num', val=70)  
            cfvo5 = FormatObject(type='num', val=85)  
            
            icon_set = IconSet(iconSet='5Arrows', cfvo=[cfvo1, cfvo2, cfvo3, cfvo4, cfvo5])
            rule = Rule(type='iconSet', iconSet=icon_set)
            
            ws.conditional_formatting.add(f"E3:E{last_student_row}", rule)
        
    first_sheet = wb.worksheets[0]
    first_sheet.title = class_name
    
    first_sheet["A1"] = f"{class_name} - {module_name}"
    current_font = first_sheet["A1"].font
    if current_font:
        first_sheet["A1"].font = Font(name=current_font.name, size=20, bold=current_font.bold, italic=current_font.italic, color=current_font.color)
    else:
        first_sheet["A1"].font = Font(size=20, bold=True)
        
    for i in range(1, len(wb.worksheets)):
        wb.worksheets[i]["A1"] = f"='{class_name}'!A1"
    
    advisor_found = False
    for row in first_sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "Advisor" in cell.value:
                cell.value = f"Advisor: {advisor_name}"
                advisor_found = True
                break
        if advisor_found:
            break
    
    start_row = 3
    for i, student in enumerate(students):
        first_sheet.cell(row=start_row + i, column=1, value=student["index"])
        first_sheet.cell(row=start_row + i, column=2, value=student["number"])
        first_sheet.cell(row=start_row + i, column=3, value=student["name"])
        first_sheet.cell(row=start_row + i, column=4, value=student["surname"])
        
    cfvo1_main = FormatObject(type='num', val=0)   
    cfvo2_main = FormatObject(type='num', val=45)  
    cfvo3_main = FormatObject(type='num', val=60)  
    cfvo4_main = FormatObject(type='num', val=70)  
    cfvo5_main = FormatObject(type='num', val=85)  
    
    icon_set_main = IconSet(iconSet='5Arrows', cfvo=[cfvo1_main, cfvo2_main, cfvo3_main, cfvo4_main, cfvo5_main])
    rule_arrows_main = Rule(type='iconSet', iconSet=icon_set_main)
    
    first_sheet.conditional_formatting.add(f"E3:E{first_sheet_last_row}", rule_arrows_main)
    first_sheet.conditional_formatting.add(f"F3:F{first_sheet_last_row}", rule_arrows_main)
    first_sheet.conditional_formatting.add(f"M3:M{first_sheet_last_row}", rule_arrows_main)

    cfvo1_L = FormatObject(type='num', val=0)       
    cfvo2_L = FormatObject(type='num', val=46.99)   
    cfvo3_L = FormatObject(type='num', val=49.99)   
    
    icon_set_L = IconSet(iconSet='3TrafficLights1', cfvo=[cfvo1_L, cfvo2_L, cfvo3_L])
    rule_icon_L = Rule(type='iconSet', iconSet=icon_set_L)
    
    first_sheet.conditional_formatting.add(f"L3:L{first_sheet_last_row}", rule_icon_L)
    
    cfvo1_N = FormatObject(type='num', val=0)       
    cfvo2_N = FormatObject(type='num', val=56.99)   
    cfvo3_N = FormatObject(type='num', val=59.5)    
    
    icon_set_N = IconSet(iconSet='3Symbols2', cfvo=[cfvo1_N, cfvo2_N, cfvo3_N])
    rule_icon_N = Rule(type='iconSet', iconSet=icon_set_N)
    
    first_sheet.conditional_formatting.add(f"N3:N{first_sheet_last_row}", rule_icon_N)
        
    white_bold_font = Font(color="FFFFFF", bold=True)
    rule_F = CellIsRule(operator='equal', formula=['"F"'], stopIfTrue=True, fill=PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid"), font=white_bold_font)
    rule_C = CellIsRule(operator='equal', formula=['"C"'], stopIfTrue=True, fill=PatternFill(start_color="4E8542", end_color="4E8542", fill_type="solid"), font=white_bold_font)
    rule_B = CellIsRule(operator='equal', formula=['"B"'], stopIfTrue=True, fill=PatternFill(start_color="1B587C", end_color="1B587C", fill_type="solid"), font=white_bold_font)
    rule_A = CellIsRule(operator='equal', formula=['"A"'], stopIfTrue=True, fill=PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid"), font=white_bold_font)
    
    letter_grade_range = f"O3:O{first_sheet_last_row}"
    first_sheet.conditional_formatting.add(letter_grade_range, rule_F)
    first_sheet.conditional_formatting.add(letter_grade_range, rule_C)
    first_sheet.conditional_formatting.add(letter_grade_range, rule_B)
    first_sheet.conditional_formatting.add(letter_grade_range, rule_A)
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

st.title("Excel Gradebook Generator")

class_lists_file = st.file_uploader("Class Lists (Excel)", type=["xlsx"])
module_name = st.text_input("Module Name (e.g., Module 3)", value="Module 3")

st.subheader("Gradebook Templates")
col1, col2 = st.columns(2)
with col1:
    a1_template = st.file_uploader("A1 Gradebook", type=["xltx", "xlsx"])
    a2_template = st.file_uploader("A2 Gradebook", type=["xltx", "xlsx"])
with col2:
    b1_template = st.file_uploader("B1 Gradebook", type=["xltx", "xlsx"])
    b2_template = st.file_uploader("B2 Gradebook", type=["xltx", "xlsx"])

if st.button("Generate Gradebooks"):
    templates = {
        "A1": a1_template,
        "A2": a2_template,
        "B1": b1_template,
        "B2": b2_template
    }
    
    if not class_lists_file:
        st.error("Lütfen Class Lists dosyasını yükleyin.")
    else:
        with st.spinner("Dosyalar oluşturuluyor..."):
            class_wb = openpyxl.load_workbook(class_lists_file, data_only=True)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for sheet_name in class_wb.sheetnames:
                    level = sheet_name.split(".")[0]
                    
                    if level in templates and templates[level]:
                        ws = class_wb[sheet_name]
                        students, advisor_name = get_class_info_from_sheet(ws)
                        
                        if not students:
                            continue
                            
                        file_data = process_class_template(templates[level].getvalue(), sheet_name, students, module_name, advisor_name)
                        zip_file.writestr(f"{level}/{sheet_name} Gradebook.xlsx", file_data)

            zip_buffer.seek(0)
            st.success("Tüm Gradebook dosyaları başarıyla oluşturuldu!")
            st.download_button(
                label="Oluşturulan Dosyaları İndir (ZIP)",
                data=zip_buffer,
                file_name="Gradebooks.zip",
                mime="application/zip"
            )
