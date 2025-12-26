import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v7.0 (Formula Clone)", layout="wide")

# --- HÜCRE KLONLAMA (EN KRİTİK FONKSİYON) ---
def clone_cell(source_cell, target_cell):
    """
    Kaynak hücredeki STİLİ ve FORMÜLÜ hedef hücreye kopyalar.
    Formülleri (örn: A5 -> A6) otomatik kaydırır.
    """
    # 1. Stil Kopyala
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # 2. Formül veya Değer Kopyala
    if source_cell.data_type == 'f':
        # Hücre formül ise (Örn: ='Main'!B5)
        try:
            # Formülü yeni satıra göre güncelle (='Main'!B6)
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            # Çevrilemezse aynısını yapıştır
            target_cell.value = source_cell.value
    else:
        # Formül değilse, şablondaki sabit bir metin olabilir (Örn: "0" veya "-")
        # Bunu da kopyalayalım ki şablon bozulmasın.
        # ANCAK: Eğer kaynak hücre boşsa kopyalama yapma.
        if source_cell.value is not None:
             target_cell.value = source_cell.value

# --- TABLO YAPISINI ÇÖZME ---
def analyze_structure(ws):
    """
    Header (Başlangıç) ve Footer (Bitiş) satırlarını tespit eder.
    """
    start_row = 6 # Varsayılan
    
    # 1. Header'ı Bul
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val or "no" in val:
                    start_row = cell.row + 1
                    break
        if start_row > 6: break
        
    # 2. Footer'ı Bul (Total/Advisor/Ortalama)
    # start_row'dan aşağı inip arıyoruz.
    current_row = start_row
    footer_row = 0
    
    # Geniş anahtar kelime havuzu
    keywords = [
        "total", "advisor", "average", "toplam", "ortalama", 
        "checker", "grade", "score", "imza", "signature", "final", "met"
    ]
    
    # Maksimum 300 satır aşağı bak
    while current_row < start_row + 300:
        # Satırın ilk 5 sütunundaki metinleri birleştirip ara
        row_text = ""
        for c in range(1, 6):
            val = ws.cell(row=current_row, column=c).value
            if val: row_text += str(val).lower()
        
        if any(k in row_text for k in keywords):
            footer_row = current_row
            break
        
        # Güvenlik: Eğer satırın kenarlığı yoksa ve boşsa, tablo bitmiş olabilir.
        # Ama şimdilik keyword araması en güvenlisi.
        
        current_row += 1
        
    if footer_row == 0:
        footer_row = start_row + 30 # Bulamazsa varsayılan
        
    return start_row, footer_row

# --- RESIZE VE POPULATE ---
def process_sheet(ws, num_students):
    start_row, footer_row = analyze_structure(ws)
    
    # Mevcut Kapasite
    current_capacity = footer_row - start_row
    
    # Hedeflenen
    needed_rows = num_students
    
    # --- DURUM A: EKLEME YAP (INSERT) ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # Ekleme Noktası: Footer'ın tam üstü.
        insert_pos = footer_row
        
        # DNA Kaynağı (Source Row): Footer'ın bir üstündeki satır (Mevcut son boş satır)
        # Bu satırda formüller ve kenarlıklar doğrudur.
        source_row_idx = footer_row - 1
        
        # 1. Satırları Ekle (Formatsız gelir)
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # 2. Kaynak Satırı Yeni Satırlara Kopyala
        max_col = ws.max_column
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source_cell = ws.cell(row=source_row_idx, column=col)
                target_cell = ws.cell(row=target_row_idx, column=col)
                clone_cell(source_cell, target_cell)
                
    # --- DURUM B: SİLME YAP (DELETE) ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye sondan başla (Footer'ın hemen üstünden yukarı doğru)
        delete_pos = start_row + needed_rows
        ws.delete_rows(delete_pos, amount=rows_to_delete)
        
    return start_row

# --- BAŞLIKLARI GÜNCELLE ---
def update_info(ws, class_name, module_name, advisor_name):
    # Sheet ismi (Sadece 1. sayfa)
    if ws.parent.index(ws) == 0:
        try: ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        except: pass

    # Smart Search
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            val = str(cell.value)
            if "GRADEBOOK" in val and "MODULE" in val:
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            if "Advisor:" in val:
                cell.value = f"Advisor: {advisor_name}"

# --- ANA İŞLEM ---
def process_workbook(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # TÜM SHEETLER İÇİN
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Başlık ve Advisor
        update_info(ws, class_name, module_name, advisor)
        
        # Boyutlandır ve Formülleri Kopyala
        data_start = process_sheet(ws, len(students_df))
        
        # SADECE MAIN SHEET (İLK SAYFA) İÇİN VERİ GİR
        # Diğer sayfalar 'process_sheet' içindeki clone_cell sayesinde
        # Main Sheet'ten formülle beslenecek.
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                
                # Sadece formül OLMAYAN hücrelere yaz (Main sheet'te isimler manueldir)
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    # KAYDETME
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker (Temizlik)
    keeps = ["MidTerm", "MET", "Midterm"]
    to_del = [s for s in wb.sheetnames if s not in keeps]
    for s in to_del: del wb[s]
    
    chk_io = None
    if len(wb.sheetnames) > 0:
        chk_io = io.BytesIO()
        wb.save(chk_io)
        chk_io.seek(0)
        
    return main_io, chk_io

# --- UI KISMI ---
st.title("🎓 Gradebook Pro v7.0 (Kesin Çözüm)")
st.markdown("""
Bu versiyon, diğer sheetlerdeki formüllerin (Örn: `='Main'!A5`) silinmemesini ve 
yeni eklenen satırlara (Örn: `='Main'!A6`) olarak kopyalanmasını garanti eder.
""")

c1, c2 = st.columns(2)
mod_in = c1.text_input("Modül İsmi", "MODULE 2")
st_file = st.file_uploader("Öğrenci Listesi (Excel)", type=["xlsx"])

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
        tmp_file = st.file_uploader("Şablon (Formülleri Açık)", type=["xlsx"])
        if tmp_file and st.button("Başlat"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.success("Tamamlandı!")
            st.download_button("Dosyaları İndir", z_buf.getvalue(), "Gradebook_Final.zip", "application/zip")
