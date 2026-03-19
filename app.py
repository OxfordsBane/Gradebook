import streamlit as st
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange
import re
import io
import zipfile

def get_students_from_sheet(sheet):
    students = []
    start_reading = False
    for row in sheet.iter_rows(values_only=True):
        if row[1] == "STUDENT NUMBER":
            start_reading = True
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
    return students

def get_current_student_rows(ws, start_row=3):
    """
    A sütunundaki (Index) sıra numaralarını sayarak şablonun o sayfasındaki 
    gerçek satır sayısını dinamik olarak bulur.
    """
    count = 0
    for r in range(start_row, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip().isdigit():
            count += 1
        else:
            break
            
    if count > 0:
        return count
        
    # Eğer A sütununda numara yoksa Excel tablosu sınırlarına bak
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row <= start_row <= max_row:
            return max_row - start_row + 1
            
    return 30 # Hiçbiri yoksa varsayılan

def adjust_template_rows_and_tables(ws, num_students):
    start_row = 3
    current_rows = get_current_student_rows(ws, start_row)
    
    action_row_idx = start_row + (current_rows // 2)
    if action_row_idx <= start_row:
        action_row_idx = start_row + 1
    
    # Satır Ekleme veya Silme
    if num_students > current_rows:
        rows_to_add = num_students - current_rows
        ws.insert_rows(action_row_idx, amount=rows_to_add)
        
        for r in range(action_row_idx, action_row_idx + rows_to_add):
            for col in range(1, ws.max_column + 1):
                source_cell = ws.cell(row=action_row_idx - 1, column=col)
                target_cell = ws.cell(row=r, column=col)
                if source_cell.has_style:
                    target_cell._style = source_cell._style

    elif num_students < current_rows:
        rows_to_delete = current_rows - num_students
        ws.delete_rows(action_row_idx, amount=rows_to_delete)

    last_student_row = start_row + num_students - 1

    # Tablo referans sınırlarını güncelleme
    for table in ws.tables.values():
        ref = table.ref
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        original_data_end = start_row + current_rows - 1
        offset = max_row - original_data_end
        if offset < 0:
            offset = 0
        new_table_max_row = last_student_row + offset
        table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_table_max_row}"

    # Formülleri uyarla
    for r in range(start_row + 1, last_student_row + 1):
        for col in range(1, ws.max_column + 1):
            master_cell = ws.cell(row=start_row, column=col)
            target_cell = ws.cell(row=r, column=col)
            
            if master_cell.data_type == 'f' and master_cell.value:
                try:
                    target_cell.value = Translator(master_cell.value, origin=master_cell.coordinate).translate_formula(target_cell.coordinate)
                except:
                    target_cell.value = master_cell.value

    # Koşullu Biçimlendirmeleri (CF) Öğrenci Satırlarında Sınırlandır (Mavi Alana Taşmasını Engelle)
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
                    if int(srow) <= start_row and int(erow) >= start_row:
                        # Kural öğrenci listesindeyse, KESİNLİKLE son öğrenci satırında bitir!
                        new_ranges.append(f"{scol}{start_row}:{ecol}{last_student_row}")
                    else:
                        new_ranges.append(rng)
                elif match_cell:
                    col, row = match_cell.groups()
                    if int(row) == start_row:
                        new_ranges.append(f"{col}{start_row}:{col}{last_student_row}")
                    else:
                        new_ranges.append(rng)
                else:
                    new_ranges.append(rng)
            
            new_sqref_str = " ".join(new_ranges)
            try:
                new_sqref = MultiCellRange(new_sqref_str)
                new_cf_rules[new_sqref] = rules
            except:
                new_cf_rules[sqref] = rules
                
        ws.conditional_formatting._cf_rules = new_cf_rules

def process_class_template(template_bytes, class_name, students, module_name):
    wb = openpyxl.load_workbook(filename=io.BytesIO(template_bytes))
    wb.template = False 
    
    # Tüm sekmelerdeki tabloları öğrenci mevcuduna göre ayarla
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        adjust_template_rows_and_tables(ws, len(students))
        
    # İLK SAYFAYA ÖZEL İŞLEMLER
    first_sheet = wb.worksheets[0]
    # 1. Sekmenin adını sınıf kodu yap
    first_sheet.title = class_name
    # 2. A1 hücresine (Başlık) Sınıf Kodu ve Modül Adını yazdır
    first_sheet["A1"] = f"{class_name} - {module_name}"
    
    # Öğrenci verilerini ilk sayfaya yazdır
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

# Modül Adı Girişi
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
                        students = get_students_from_sheet(ws)
                        
                        if not students:
                            continue
                            
                        # process_class_template fonksiyonuna module_name değişkenini de gönderiyoruz
                        file_data = process_class_template(templates[level].getvalue(), sheet_name, students, module_name)
                        zip_file.writestr(f"{level}/{sheet_name} Gradebook.xlsx", file_data)

            zip_buffer.seek(0)
            st.success("Tüm Gradebook dosyaları başarıyla oluşturuldu!")
            st.download_button(
                label="Oluşturulan Dosyaları İndir (ZIP)",
                data=zip_buffer,
                file_name="Gradebooks.zip",
                mime="application/zip"
            )
