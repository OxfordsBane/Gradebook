import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import range_boundaries, get_column_letter
from copy import copy
import io
import zipfile

# --- SAYFA AYARI ---
st.set_page_config(page_title="v13.0 - TABLE & FORMAT FIX", layout="wide")

# --- HÜCRE KOPYALAMA ---
def clone_cell(source_cell, target_cell):
    """Stil, Formül ve Değer Kopyalar"""
    # 1. Stil
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # 2. Formül
    if source_cell.data_type == 'f':
        try:
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            target_cell.value = source_cell.value
    # 3. Sabit Değer (Boş değilse)
    elif source_cell.value is not None:
        target_cell.value = source_cell.value

# --- TABLO NESNESİNİ (KÖŞEDEKİ ÇUBUĞU) GÜNCELLEME ---
def fix_table_boundaries(ws, rows_added):
    """
    Sayfadaki Excel Tablolarını (ListObjects) bulur.
    Bitiş satırını (rows_added) kadar aşağı öteleyerek günceller.
    """
    if not ws.tables:
        return

    for table in ws.tables.values():
        # Tablonun mevcut sınırlarını çöz (A1:F30 gibi)
        min_c, min_r, max_c, max_r = range_boundaries(table.ref)
        
        # Yeni bitiş satırını hesapla
        new_max_r = max_r + rows_added
        
        # Yeni koordinatı oluştur (Örn: A1:F90)
        new_ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{new_max_r}"
        
        # Tabloya yeni sınırını ata
        table.ref = new_ref

# --- HEADER BULMA ---
def find_header_row(ws):
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val:
                    return cell.row
    return 6

# --- SHEET İŞLEME MOTORU ---
def process_sheet_resize(ws, num_students):
    header_row = find_header_row(ws)
    
    # GÜVENLİ BÖLGE: Header'ın 5 satır altı (Tablonun ortası)
    insert_pos = header_row + 5
    
    # Varsayılan Kapasite (Şablon Genelde 30'dur)
    current_capacity = 30 
    
    # Kapasiteyi Dinamik Kontrol Et (Opsiyonel ama güvenli)
    # Tablonun çizgilerinin nerede bittiğine bakarak da anlayabiliriz
    # Ama standart 30 varsayımı sizin şablonlarınızda tutuyor.
    
    needed_rows = num_students
    rows_added = 0 

    # --- DURUM A: EKLEME YAP ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        rows_added = rows_to_add
        
        # 1. Satırları Araya Ekle
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # 2. Üstteki Satırı Aşağı Kopyala (Fill Down)
        ref_row_idx = insert_pos - 1
        max_col = ws.max_column
        
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row_idx, column=col)
                target = ws.cell(row=target_row_idx, column=col)
                clone_cell(source, target)

    # --- DURUM B: SİLME YAP ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        rows_added = -rows_to_delete
        ws.delete_rows(insert_pos, amount=rows_to_delete)
        
    # --- KRİTİK: TABLO NESNESİNİ GÜNCELLE ---
    # Eğer satır eklediysek, tablonun 'ref' özelliğini güncellemeliyiz.
    if rows_added != 0:
        fix_table_boundaries(ws, rows_added)
        
    return header_row + 1

# --- BAŞLIKLARI GÜNCELLE ---
def update_headers(ws, class_name, module_name, advisor_name):
    # Sheet Adı
    if ws.parent.index(ws) == 0:
        try: ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        except: pass

    # İçerik Güncelleme
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            val = str(cell.value)
            if "GRADEBOOK" in val and "MODULE" in val:
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            if "Advisor:" in val:
                cell.value = f"Advisor: {advisor_name}"

# --- MAIN ---
def process_workbook_v13(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Başlıkları güncelle
        update_headers(ws, class_name, module_name, advisor)
        
        # Resize ve Tablo Fix
        data_start = process_sheet_resize(ws, len(students_df))
        
        # Main Sheet Veri Girişi
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
st.title("🛠️ Gradebook v13.0 - FINAL TABLE FIX")
st.error("DİKKAT: Eski kodu durdurup bunu yeniden başlattığınızdan emin olun.")

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
        tmp_file = st.file_uploader("Master Şablon", type=["xlsx"])
        if tmp_file and st.button("OLUŞTUR"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v13(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.success("Tüm tablolar yeniden boyutlandırıldı!")
            st.download_button("İndir", z_buf.getvalue(), "Gradebook_v13.zip", "application/zip")
