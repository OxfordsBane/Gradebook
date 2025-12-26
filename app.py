import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v5.0 (Final Clone)", layout="wide")

# --- HÜCRE KOPYALAMA FONKSİYONU ---
def clone_cell(source_cell, target_cell):
    """
    Bir hücrenin stilini ve formülünü hedef hücreye kopyalar.
    """
    # 1. STİL KOPYALA
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # 2. DEĞER/FORMÜL KOPYALA
    if source_cell.data_type == 'f':
        # Eğer hücrede formül varsa, onu yeni satıra göre ötele (Translate)
        # Örn: A5 -> A6 olur.
        try:
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            # Çevrilemezse (örn: sabit isimli range) aynısını yaz
            target_cell.value = source_cell.value
    else:
        # Formül yoksa değeri kopyalamıyoruz (Boş gelmesi daha güvenli)
        # Ancak şablondaki sabit yazıları (varsa) korumak isterseniz burayı açabiliriz.
        # Şimdilik sadece stil ve formül odaklıyız.
        pass

# --- TABLO ALANINI BULMA ---
def find_data_zone(ws):
    """
    Şablonun nerede başlayıp nerede bittiğini bulur.
    Footer (Total/Advisor) kısmını bozmamak için bitişin hemen öncesine ekleme yapacağız.
    """
    start_row = 6 # Varsayılan güvenlik
    
    # 1. Header'ı Bul
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val:
                    start_row = cell.row + 1
                    break
        if start_row > 10: break # Çok aşağı indiyse dur
        
    # 2. Footer'ı Bul (Total/Advisor)
    # start_row'dan itibaren aşağı inip border/yazı kontrolü yapıyoruz.
    current_row = start_row
    max_search = 300
    footer_row = start_row + 29 # Bulamazsa varsayılan
    
    while current_row < start_row + max_search:
        # A ve B sütunundaki değerleri kontrol et
        c1 = ws.cell(row=current_row, column=1).value
        c2 = ws.cell(row=current_row, column=2).value
        val_str = (str(c1) + str(c2)).lower()
        
        keywords = ["total", "advisor", "average", "toplam", "ortalama", "checker"]
        if any(k in val_str for k in keywords):
            footer_row = current_row
            break
        current_row += 1
        
    return start_row, footer_row

# --- SHEET DÜZENLEME ---
def process_sheet_resize(ws, num_students):
    """
    Her sheet için satır ekleme/silme işlemini yapar.
    """
    start_row, footer_row = find_data_zone(ws)
    
    # Mevcut kapasite (Footer ile Header arasındaki boş satırlar)
    current_capacity = footer_row - start_row
    
    # Hedeflenen satır sayısı
    needed_rows = num_students
    
    # --- DURUM 1: SATIR EKLEME (INSERT) ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # Ekleme noktası: Footer'ın hemen üstü (Mevcut son boş satırın altı)
        insert_pos = footer_row
        
        # Satırları ekle (Excel bunları formatsız ekler)
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Referans Satırı: Ekleme yaptığımız yerin BİR ÜSTÜNDEKİ satır.
        # Bu satırın formatını ve formülünü aşağıya doğru kopyalayacağız (Fill Down).
        ref_row_idx = insert_pos - 1
        max_col = ws.max_column
        
        # Eklenen her yeni satır için döngü
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            
            for col in range(1, max_col + 1):
                source_cell = ws.cell(row=ref_row_idx, column=col)
                target_cell = ws.cell(row=target_row_idx, column=col)
                
                clone_cell(source_cell, target_cell)
                
    # --- DURUM 2: SATIR SİLME (DELETE) ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye sondan başla (Footer'ın üstünden)
        delete_pos = start_row + needed_rows
        ws.delete_rows(delete_pos, amount=rows_to_delete)
        
    return start_row

# --- BAŞLIK GÜNCELLEME ---
def update_headers(ws, class_name, module_name, advisor_name):
    # Sadece ilk sheetin ismini değiştir
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

# --- ANA MOTOR ---
def process_workbook_final(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # 1. TÜM SHEETLERİ DÖNGÜYE AL
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Başlıkları güncelle
        update_headers(ws, class_name, module_name, advisor)
        
        # Sheet'i boyutlandır (Ekle/Sil ve Formülleri Kopyala)
        data_start_row = process_sheet_resize(ws, len(students_df))
        
        # 2. VERİ GİRİŞİ (SADECE MAIN SHEET)
        # Sadece Workbook'un ilk sayfasına (Main Sheet) isimleri yazıyoruz.
        # Diğer sayfalar process_sheet_resize içindeki clone_cell fonksiyonu sayesinde
        # Main Sheet'e bağlı formülleri kopyaladığı için otomatik dolacak.
        
        if wb.index(ws) == 0: 
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start_row + i
                
                # Formül olmayan hücrelere veri yaz
                # (Main Sheet'te isimler manuel girilir, formül değildir)
                
                # Index
                c1 = ws.cell(r, 1)
                if c1.data_type != 'f': c1.value = i + 1
                
                # No
                c2 = ws.cell(r, 2)
                if c2.data_type != 'f': c2.value = student[col_map['no']]
                
                # Ad
                c3 = ws.cell(r, 3)
                if c3.data_type != 'f': c3.value = student[col_map['name']]
                
                # Soyad
                c4 = ws.cell(r, 4)
                if c4.data_type != 'f': c4.value = student[col_map['surname']]

    # KAYDET
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker Dosyaları
    sheets_keep = ["MidTerm", "MET", "Midterm"]
    to_del = [s for s in wb.sheetnames if s not in sheets_keep]
    for s in to_del: del wb[s]
    
    chk_io = None
    if len(wb.sheetnames) > 0:
        chk_io = io.BytesIO()
        wb.save(chk_io)
        chk_io.seek(0)
        
    return main_io, chk_io

# --- UI ---
st.title("🎓 Gradebook Pro v5.0 (Final)")
st.markdown("Satırları kopyalarken formülleri ve stilleri de kopyalar.")

col1, col2 = st.columns(2)
mod_in = col1.text_input("Modül", "MODULE 2")
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
                    m, ch = process_workbook_final(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(cls_list))
            
            st.success("Bitti!")
            st.download_button("İndir", z_buf.getvalue(), "Gradebook_v5.zip", "application/zip")
