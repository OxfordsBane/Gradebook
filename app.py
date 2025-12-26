import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.cell_range import CellRange
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook v11.0 (Table Resize)", layout="wide")

# --- 1. HÜCRE KLONLAMA (STİL + FORMÜL) ---
def clone_cell(source_cell, target_cell):
    """v9.0'daki stabil kopyalama mantığı."""
    # Stil
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # Formül
    if source_cell.data_type == 'f':
        try:
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            target_cell.value = source_cell.value
    elif source_cell.value is not None:
        target_cell.value = source_cell.value

# --- 2. EXCEL TABLOSUNU (KÖŞEDEKİ ÇUBUĞU) UZATMA ---
def extend_excel_table(ws, rows_added):
    """
    Sayfadaki 'Table' nesnesini bulur ve sınırını (Range) 
    eklenen satır kadar aşağı çeker.
    """
    if not ws.tables:
        return

    for table in ws.tables.values():
        # Tablonun mevcut aralığını oku (Örn: "A5:G35")
        ref = table.ref
        cr = CellRange(ref)
        
        # Tablonun bittiği satırı güncelle
        # (Mevcut Bitiş + Eklenen Satır Sayısı)
        new_max_row = cr.max_row + rows_added
        
        # Yeni aralığı oluştur (Örn: "A5:G95")
        cr.max_row = new_max_row
        new_ref = cr.coord
        
        # Tabloya yeni sınırını bildir
        table.ref = new_ref

# --- 3. TABLO BAŞLANGICINI BUL (v9.0 MANTIĞI) ---
def find_header_row(ws):
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val:
                    return cell.row
    return 6

# --- 4. SHEET İŞLEME VE RESIZE ---
def process_sheet_resize(ws, num_students):
    header_row = find_header_row(ws)
    
    # GÜVENLİ BÖLGE: Header'ın 5 satır altı (Tablonun göbeği)
    insert_pos = header_row + 5
    
    # Şablondaki varsayılan kapasite (Genelde 30)
    current_capacity = 30 
    needed_rows = num_students
    
    rows_added_count = 0 # Tabloyu ne kadar uzatacağımızı tutalım

    # DURUM A: EKLEME
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        rows_added_count = rows_to_add
        
        # 1. Satırları Araya Ekle
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # 2. Üst Satırı Aşağı Kopyala (Fill Down)
        ref_row_idx = insert_pos - 1
        max_col = ws.max_column
        
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row_idx, column=col)
                target = ws.cell(row=target_row_idx, column=col)
                clone_cell(source, target)

    # DURUM B: SİLME
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        rows_added_count = -rows_to_delete # Negatif büyüme
        ws.delete_rows(insert_pos, amount=rows_to_delete)
        
    # --- KRİTİK ADIM: TABLO NESNESİNİ GÜNCELLE ---
    # Hücreleri boyadık ama Excel'e "Tablo uzadı" dememiz lazım.
    if rows_added_count != 0:
        extend_excel_table(ws, rows_added_count)
        
    return header_row + 1

# --- 5. BAŞLIK GÜNCELLEME ---
def update_headers(ws, class_name, module_name, advisor_name):
    if ws.parent.index(ws) == 0:
        try: ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        except: pass

    for row in ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            val = str(cell.value)
            if "GRADEBOOK" in val and "MODULE" in val:
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            if "Advisor:" in val:
                cell.value = f"Advisor: {advisor_name}"

# --- 6. ANA SÜREÇ ---
def process_workbook_v11(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        update_headers(ws, class_name, module_name, advisor)
        
        # Resize + Tablo Nesnesi Güncelleme
        data_start = process_sheet_resize(ws, len(students_df))
        
        # Sadece Main Sheet Veri Girişi
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    keeps = ["MidTerm", "MET", "Midterm"]
    dels = [s for s in wb.sheetnames if s not in keeps]
    for s in dels: del wb[s]
    
    chk_io = None
    if len(wb.sheetnames) > 0:
        chk_io = io.BytesIO()
        wb.save(chk_io)
        chk_io.seek(0)
        
    return main_io, chk_io

# --- UI ---
st.title("🎓 Gradebook v11.0 (Table Object Resize)")
st.markdown("""
**v9.0 Mantığı + Tablo Nesnesi Güncelleme:**
Bu sürüm, hücreleri kopyaladıktan sonra Excel'in **"Tablo Çubuğunu" (Table Handle)** otomatik olarak aşağı çeker.
Böylece tablo görsel olarak da formül olarak da son öğrenciye kadar uzar.
""")

c1, c2 = st.columns(2)
mod_in = c1.text_input("Modül İsmi", "MODULE 2")
st_file = st.file_uploader("Öğrenci Listesi", type=["xlsx"])

if st_file:
    df = pd.read_excel(st_file)
    cols = st.columns(5)
    col_map = {
        'class': cols[0].selectbox("Sınıf", df.columns, 0),
        'no': cols[1].selectbox("No", df.columns, 1),
        'name': cols[2].selectbox("Ad", df.columns, 2),
        'surname': cols[3].selectbox("Soyad", df.columns, 3),
        'advisor': cols[4].selectbox("Advisor", df.columns, 4 if len(df.columns)>4 else 0)
    }
    
    classes = st.multiselect("Sınıflar", df[col_map['class']].unique())
    if classes:
        tmp_file = st.file_uploader("Şablon", type=["xlsx"])
        if tmp_file and st.button("OLUŞTUR"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v11(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.success("İşlem Tamam!")
            st.download_button("İndir", z_buf.getvalue(), "Gradebook_v11.zip", "application/zip")
