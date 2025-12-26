import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v6.0 (Safe Zone Fix)", layout="wide")

# --- HÜCRE KLONLAMA (STİL + FORMÜL) ---
def clone_cell(source_cell, target_cell):
    """Bir hücrenin tüm genetiğini (stil, formül, kilit) kopyalar."""
    # 1. Stil
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # 2. Formül (Translate)
    if source_cell.data_type == 'f':
        try:
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            target_cell.value = source_cell.value

# --- TABLO SINIRLARINI BULMA (GELİŞMİŞ) ---
def find_table_boundaries(ws):
    """
    Tablonun başlangıcını (Header) ve bitişini (Footer/Bottom)
    garantili bir şekilde bulmaya çalışır.
    """
    start_row = 6 # Varsayılan güvenlik
    
    # 1. Header'ı Bul (Index, No, Student)
    header_found = False
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val:
                    start_row = cell.row + 1 # Veri, başlığın altından başlar
                    header_found = True
                    break
        if header_found: break
        
    # 2. Footer'ı Bul (Bitiş Noktası)
    # start_row'dan itibaren aşağı inip "Tablo nerede bitiyor?" diye bakacağız.
    # Strateji: "Advisor/Total" bulursak orasıdır. Bulamazsak dolu olan son satırdır.
    
    current_row = start_row
    max_search = 300
    footer_row = start_row + 30 # Hiçbir şey bulamazsak varsayılan 30 satır
    found_keyword = False
    
    # Anahtar kelimeler
    keywords = ["total", "advisor", "average", "toplam", "ortalama", "checker", "grade", "score", "imza", "signature"]
    
    while current_row < start_row + max_search:
        # A, B, C sütunlarına bak (Genelde yazılar buradadır)
        val_str = ""
        for c in range(1, 5):
            val = ws.cell(row=current_row, column=c).value
            val_str += str(val).lower() if val else ""
            
        # 1. Kriter: Kelime Eşleşmesi
        if any(k in val_str for k in keywords):
            footer_row = current_row
            found_keyword = True
            break
            
        current_row += 1
        
    # Eğer kelime bulamadıysak (Örn: Role-play sheet'inde total yazmıyorsa)
    # Görsel/Dolu Satır kontrolü yapalım:
    if not found_keyword:
        # Tersten yukarı çıkalım (max_row'dan geriye)
        # Ama tüm sheet dolu olabilir, o yüzden start_row'dan aşağı inip
        # "Art arda 5 tane tamamen boş ve kenarlıksız satır" görünce duralım.
        
        check_row = start_row
        empty_streak = 0
        last_data_row = start_row
        
        while check_row < start_row + 100:
            is_empty = True
            # Satırın ilk 10 hücresine bak
            for c in range(1, 11):
                cell = ws.cell(row=check_row, column=c)
                if cell.value or (cell.border and (cell.border.top.style or cell.border.bottom.style)):
                    is_empty = False
                    break
            
            if is_empty:
                empty_streak += 1
            else:
                empty_streak = 0
                last_data_row = check_row
                
            if empty_streak > 5: # 5 satır boşluk varsa tablo bitmiştir
                break
            check_row += 1
            
        footer_row = last_data_row + 1

    # Güvenlik Kontrolü: Footer header'dan çok yakınsa (hatalıysa) düzelt
    if footer_row <= start_row + 1:
        footer_row = start_row + 30

    return start_row, footer_row

# --- SHEET DÜZENLEME ---
def process_sheet_resize(ws, num_students):
    start_row, footer_row = find_table_boundaries(ws)
    
    # Mevcut Kapasite (Footer ile Header arası)
    current_capacity = footer_row - start_row
    needed_rows = num_students
    
    # --- DURUM 1: EKLEME YAP (INSERT) ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # KRİTİK NOKTA: Footer'ın tam üstüne değil, "1 satır üstüne" ekleyelim.
        # Böylece footer ile veri arasına girmemiş oluruz, footer'ı aşağı iteriz.
        # Sizin "25. satır" taktiği.
        
        insert_pos = footer_row - 1 
        
        # Eğer tablo çok küçükse (1-2 satırsa) hata olmasın
        if insert_pos < start_row: insert_pos = start_row
        
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Stil Referansı: Insert yaptığımız yerin hemen üstündeki satır
        ref_row_idx = insert_pos - 1
        # Eğer üst satır header ise (tablo boşsa), mecburen insert_pos'un kendisini (yeni boş satır) değil, start_row'u referans al
        if ref_row_idx < start_row: ref_row_idx = start_row
        
        # Kopyalama Döngüsü
        max_col = ws.max_column
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row_idx, column=col)
                target = ws.cell(row=target_row_idx, column=col)
                clone_cell(source, target)
                
    # --- DURUM 2: SİLME YAP (DELETE) ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye sondan başla (Footer'ın hemen üstünden yukarı doğru)
        # delete_start = footer_row - rows_to_delete
        # Daha güvenli: Verilerin bittiği yerden başla
        delete_pos = start_row + needed_rows
        ws.delete_rows(delete_pos, amount=rows_to_delete)
        
    return start_row

# --- BAŞLIK GÜNCELLEME ---
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

# --- ANA İŞLEM ---
def process_workbook_v6(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # Tüm Sheetleri İşle
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        update_headers(ws, class_name, module_name, advisor)
        
        # Resize ve Formül Kopyalama
        data_start = process_sheet_resize(ws, len(students_df))
        
        # Sadece Main Sheet'e İsim Yaz
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                # Formül olmayan hücrelere veri bas
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    # Kaydet
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker
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
st.title("🎓 Gradebook Pro v6.0 (Safe Zone)")
st.markdown("Footer (Advisor/Total) satırını koruyarak, araya güvenli ekleme yapar.")

c1, c2 = st.columns(2)
mod_in = c1.text_input("Modül", "MODULE 2")
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
    
    cls_list = st.multiselect("Sınıflar", df[col_map['class']].unique())
    if cls_list:
        tmp_file = st.file_uploader("Master Şablon", type=["xlsx"])
        if tmp_file and st.button("Başlat"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(cls_list):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v6(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(cls_list))
            
            st.success("İşlem Tamam!")
            st.download_button("İndir", z_buf.getvalue(), "Gradebook_v6.zip", "application/zip")
