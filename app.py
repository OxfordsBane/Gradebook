import streamlit as st
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.styles import Font
from openpyxl.worksheet.cell_range import MultiCellRange
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

def get_current_student_rows(ws, start_row=3):
    count = 0
    for r in range(start_row, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip().isdigit():
            count += 1
        else:
            break
            
    if count > 0:
        return count
        
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row <= start_row <= max_row:
            return max_row - start_row + 1
            
    return 30

def shift_formula_rows(formula_str, threshold_row, offset):
    """
    Formül metinlerindeki hücre referanslarını (Örn: E$33, A15) tespit eder
    ve eğer eşik değerinden büyükse (yani mavi alandaysa) satır sayısını offset kadar kaydırır.
    """
    if not formula_str or not isinstance(formula_str, str) or not formula_str.startswith('='):
        return formula_str
        
    def repl(match):
        prefix = match.group(1)
        row_num = int(match.group(2))
        if row_num >= threshold_row:
            return f"{prefix}{row_num + offset}"
        return match.group(0)
        
    # Sadece geçerli hücre referanslarını eşleştirir (Örn: E33, E$33, $E$33)
    return re.sub(r'(?<![A-Za-z])(\$?[A-Za-z]{1,3}\$?)(\d+)\b', repl, formula_str)

def adjust_template_rows_and_tables(ws, num_students):
    start_row = 3
    current_rows = get_current_student_rows(ws, start_row)
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
                source_cell = ws.cell(row=action_row_idx - 1, column=col)
                target_cell = ws.cell(row=r, column=col)
                if source_cell.has_style:
                    target_cell._style = source_cell._style

    elif num_students < current_rows:
        rows_to_delete = current_rows - num_students
        ws.delete_rows(action_row_idx, amount=rows_to_delete)
        offset = -rows_to_delete

    last_student_row = start_row + num_students - 1

    # Formüllerdeki $ işaretli mutlak referansları ve mavi alan formüllerini topluca güncelle
    if offset != 0:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value:
                    cell.value = shift_formula_rows(str(cell.value), action_row_idx, offset)

    # Excel Tablo referans sınırlarını güncelle
    for table in ws.tables.values():
        ref = table.ref
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        table_offset = max_row - original_last_student_row
        if table_offset < 0:
            table_offset = 0
        new_table_max_row = last_student_row + table_offset
        table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_table_max_row}"

    # Düzeltilmiş (Usta) Formülleri tüm öğrencilere uyarlayarak kopyala
    for r in range(start_row + 1, last_student_row + 1):
        for col in range(1, ws.max_column + 1):
            master_cell = ws.cell(row=start_row, column=col)
            target_cell = ws.cell(row=r, column=col)
            
            if master_cell.data_type == 'f' and master_cell.value:
                try:
                    target_cell.value = Translator(master_cell.value, origin=master_cell.coordinate).translate_formula(target_cell.coordinate)
                except:
                    target_cell.value = master_cell.value

    # Koşullu Biçimlendirmeleri (CF) Mavi Alandan Tamamen Temizleme İşlemi
    if hasattr(ws.conditional_formatting, '_cf_rules'):
        new_cf_rules = {}
        for sqref, rules in ws.conditional_formatting._cf_rules.items():
            if hasattr(sqref, 'ranges'):
                sqref_str = " ".join([rng.coord for rng in sqref.ranges])
            else:
                sqref_str = str(sqref)
            
            sqref_str = sqref_str.replace("<MultiCellRange [", "").replace("]>", "")
            
            new_ranges = []
            for rng in sqref_str.split():
                match_range = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", rng)
                match_cell = re.match(r"^([A-Z]+)(\d+)$", rng)
                
                if match_range:
                    scol, srow, ecol, erow = match_range.groups()
                    srow_int, erow_int = int(srow), int(erow)
                    
                    if srow_int > original_last_student_row:
                        continue 
                        
                    if srow_int <= start_row and erow_int >= start_row:
                        new_ranges.append(f"{scol}{start_row}:{ecol}{last_student_row}")
                    else:
                        new_ranges.append(rng)
                        
                elif match_cell:
                    col, row = match_cell.groups()
                    row_int = int(row)
                    
                    if row_int > original_last_student_row:
                        continue
                        
                    if row_int == start_row:
                        new_ranges.append(f"{col}{start_row}:{col}{last_student_row}")
                    else:
                        new_ranges.append(rng)
                else:
                    new_ranges.append(rng)
            
            if new_ranges:
                new_sqref_str = " ".join(new_ranges)
                try:
                    new_sqref = MultiCellRange(new_sqref_str)
                    new_cf_rules[new_sqref] = rules
                except:
                    new_cf_rules[sqref] = rules
                
        ws.conditional_formatting._cf_rules = new_cf_rules

def process_class_template(template_bytes, class_name, students, module_name, advisor_name):
    wb = openpyxl.load_workbook(filename=io.BytesIO(template_bytes))
    wb.template = False 
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        adjust_template_rows_and_tables(ws, len(students))
        
    first_sheet = wb.worksheets[0]
    first_sheet.title = class_name
    
    first_sheet["A1"] = f"{class_name} - {module_name}"
    current_font = first_sheet["A1"].font
    if current_font:
        first_sheet["A1"].font = Font(name=current_font.name, size=20, bold=current_font.bold, italic=current_font.italic, color=current_font.color)
    else:
        first_sheet["A1"].font = Font(size=20, bold=True)
    
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
