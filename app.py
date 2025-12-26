import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook v10.0 (Fill Down Logic)", layout="wide")

# --- HÜCRE KOPYALAMA (FORMÜL KAYDIRMA DAHİL) ---
def replicate_row(ws, source_row_idx, target_row_idx):
    """
    Kaynak satırın tüm hücrelerini hedef satıra kopyalar.
    Formülleri (A5 -> A6) otomatik günceller.
    """
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        source_cell = ws.cell(row=source_row_idx, column=col)
        target_cell = ws.cell(row=target_row_idx, column=col)
        
        # 1. STİL (Kenarlık, Renk, Font, Kilit)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)
        
        # 2. DEĞER / FORMÜL
        if source_cell.data_type == 'f':
            # Formül ise kaydırarak kopyala
            try:
                target_cell.value = Translator(
                    source_cell.value, source_cell.coordinate
                ).translate_formula(target_cell.coordinate)
            except:
                target_cell.value = source_cell.value
        else:
            # Sabit değer ise (ve boş değilse) kopyala. 
            # (Örn: "-" işareti veya "0" puanı varsayılan olarak varsa)
            if source_cell.value is not None:
                target_cell.value = source_cell.value

# --- TABLO SINIRLARINI BULMA ---
def find_table_structure(ws):
    """
    Header (Başlık) ve Footer (Bitiş) satırlarını tespit eder.
    """
    start_row = 6 # Varsayılan
    
    # 1. Header'ı Bul
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val:
                    start_row = cell.row + 1
                    break
        if start_row > 6: break
        
    # 2. Footer'ı Bul (Tablonun bittiği yer)
    # Advisor, Total, Average gibi kelimeleri arar.
    current_row = start_row
    footer_row = 0
    keywords = ["total", "advisor", "average", "toplam", "ortalama", "checker", "grade", "score"]
    
    # 500 satıra kadar tara
    while current_row < start_row + 500:
        row_txt = ""
        # İlk 5 sütuna bakmak yeterli
        for c in range(1, 6):
            v = ws.cell(row=current_row, column=c).value
            if v: row_txt += str(v).lower()
            
        if any(k in row_txt for k in keywords):
            footer_row = current_row
            break
        
        # Görsel Kontrol: Eğer satır tamamen boşsa ve kenarlığı yoksa?
        # Bu riskli olabilir, o yüzden keyword en güvenlisi.
        # Eğer keyword yoksa, varsayılan bir sınır belirleriz.
        
        current_row += 1
    
    # Eğer footer bulunamazsa, dolu olan son satırın 2 altını footer kabul et
    if footer_row == 0:
        footer_row = ws.max_row + 1
        
    return start_row, footer_row

# --- SHEET YENİDEN BOYUTLANDIRMA ---
def resize_and_fill(ws, num_students):
    start_row, footer_row = find_table_structure(ws)
    
    # Mevcut Kapasite (Footer ile Veri Başlangıcı Arası)
    current_capacity = footer_row - start_row
    needed_rows = num_students
    
    # --- DURUM 1: KAPASİTE YETERSİZ -> EKLEME YAP ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # Ekleme Noktası: Footer'ın tam olduğu yer.
        # Excel'de "Total" satırına sağ tıklayıp "Ekle" demek gibidir.
        insert_pos = footer_row
        
        # Satırları aç (Footer aşağı kayar)
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Referans Satırı: Footer'ın hemen üstündeki satır (insert_pos - 1)
        # Bu satırda formüller ve kenarlıklar doğrudur.
        source_row = insert_pos - 1
        
        # Yeni açılan satırları (insert_pos'tan itibaren) doldur
        for i in range(rows_to_add):
            target_row = insert_pos + i
            replicate_row(ws, source_row, target_row)
            
    # --- DURUM 2: KAPASİTE FAZLA -> SİLME YAP ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye sondan başla (Footer'ın hemen üstünden yukarı doğru)
        delete_pos = start_row + needed_rows
        ws.delete_rows(delete_pos, amount=rows_to_delete)
        
    return start_row

# --- BAŞLIKLARI GÜNCELLE ---
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
def process_workbook_v10(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # TÜM SHEETLERİ GEZ
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 1. Başlıkları Güncelle
        update_headers(ws, class_name, module_name, advisor)
        
        # 2. Resize Yap ve Formülleri Uzat
        # Bu fonksiyon "Fill Down" işlemini yapar.
        data_start = resize_and_fill(ws, len(students_df))
        
        # 3. VERİ GİRİŞİ (SADECE MAIN SHEET)
        # Sadece ilk sayfaya (Main) isimleri yazıyoruz.
        # Diğer sayfalar, resize_and_fill içindeki replicate_row fonksiyonu 
        # sayesinde Main'e bağlı formülleri kopyaladığı için otomatik dolacak.
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                
                # Sadece formül olmayan hücrelere yaz
                # (Main Sheet'te isim/no sütunları genelde manueldir)
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    # KAYDET
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker Dosyaları (Temizlik)
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
st.title("🎓 Gradebook v10.0 (Fill-Down Edition)")
st.markdown("""
**Çalışma Prensibi:**
Program, tablonun son satırını tespit eder. Eğer yer yoksa, "Total" satırını aşağı iterek yer açar ve 
**bir üstteki satırın formüllerini** yeni açılan yere kopyalar (Excel'deki Fill Down / Ctrl+D işlemi).
""")

c1, c2 = st.columns(2)
mod_in = c1.text_input("Modül İsmi", "MODULE 2")
st_file = st.file_uploader("Öğrenci Listesi", type=["xlsx"])

if st_file:
    df = pd.read_excel(st_file)
    st.info(f"{len(df)} öğrenci bulundu.")
    
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
        if tmp_file and st.button("OLUŞTUR", type="primary"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v10(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.success("Bitti! Tüm formüller kopyalandı.")
            st.download_button("Dosyaları İndir", z_buf.getvalue(), "Gradebook_v10.zip", "application/zip")
