import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

# --- CACHE TEMİZLİĞİ VE AYARLAR ---
st.set_page_config(page_title="Gradebook v9.0 FINAL", layout="wide")

# --- HÜCRE KLONLAMA (DNA KOPYALAMA) ---
def clone_cell(source_cell, target_cell):
    """Stil, Formül, Kenarlık ve Kilit bilgisini kopyalar."""
    # 1. Stil
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    # 2. Formül veya Değer
    if source_cell.data_type == 'f':
        try:
            # Formülü kaydır (A5 -> A6)
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            target_cell.value = source_cell.value
    elif source_cell.value is not None:
        # Formül değilse ve boş değilse kopyala (Sabit metinler için)
        target_cell.value = source_cell.value

# --- HEADER BULUCU ---
def find_header_row(ws, debug_log):
    """Tablonun başlık satırını bulur."""
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val or "no" in val:
                    return cell.row
    return 6 # Bulamazsa varsayılan

# --- SHEET İŞLEME MOTORU ---
def process_sheet_resize(ws, num_students, debug_log):
    header_row = find_header_row(ws, debug_log)
    
    # --- STRATEJİ: GÖBEKTEN EKLEME ---
    # Tablonun header'ından 5 satır aşağısı (örn: 11. satır) her zaman güvenlidir.
    # Footer nerede olursa olsun, araya girdiğimiz için aşağı itilir.
    insert_pos = header_row + 5
    
    # Şablondaki varsayılan boş satır sayısı (Genelde 30)
    # Bunu dinamik olarak footer'ı arayarak bulmak riskliydi, o yüzden sabit varsayıyoruz
    # veya dolu satır kontrolü yapıyoruz.
    current_capacity = 30 
    
    # Hedeflenen satır
    needed_rows = num_students
    
    if debug_log:
        st.write(f"Sheet: {ws.title} | Header: {header_row} | Insert Pos: {insert_pos} | Needed: {needed_rows}")

    # DURUM A: EKLEME (INSERT)
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Referans: Ekleme yapılan yerin bir üstü
        ref_row_idx = insert_pos - 1
        
        max_col = ws.max_column
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row_idx, column=col)
                target = ws.cell(row=target_row_idx, column=col)
                clone_cell(source, target)

    # DURUM B: SİLME (DELETE)
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        ws.delete_rows(insert_pos, amount=rows_to_delete)
        
    return header_row + 1 # Veri giriş başlangıcı

# --- BAŞLIKLARI GÜNCELLE ---
def update_headers(ws, class_name, module_name, advisor_name):
    # Sheet ismi (Sadece 1. sayfa)
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

# --- ANA SÜREÇ ---
def process_workbook_v9(template_bytes, class_name, students_df, col_map, module_name, debug_log):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # TÜM SHEETLERİ DÖNGÜYE AL
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        update_headers(ws, class_name, module_name, advisor)
        
        # 1. Tabloyu Boyutlandır ve Formülleri Kopyala
        data_start = process_sheet_resize(ws, len(students_df), debug_log)
        
        # 2. SADECE MAIN SHEET'E İSİM YAZ
        # Diğer sheetler formülle çekecek
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                # Formül olmayan hücrelere yaz
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
st.title("🚀 Gradebook v9.0 FINAL EDITION")
st.markdown("""
**Dikkat:** Eğer bu başlığı görmüyorsanız kod güncellenmemiştir. Lütfen uygulamayı durdurup tekrar başlatın.
""")

debug_mode = st.checkbox("🛠️ Debug Modu (İşlem detaylarını göster)")

c1, c2 = st.columns(2)
mod_in = c1.text_input("Modül İsmi", "MODULE 2")
st_file = st.file_uploader("Öğrenci Listesi (Excel)", type=["xlsx"])

if st_file:
    df = pd.read_excel(st_file)
    st.success(f"Liste yüklendi: {len(df)} öğrenci bulundu.")
    
    cols = st.columns(5)
    col_map = {
        'class': cols[0].selectbox("Sınıf Sütunu", df.columns, 0),
        'no': cols[1].selectbox("Numara Sütunu", df.columns, 1),
        'name': cols[2].selectbox("Ad Sütunu", df.columns, 2),
        'surname': cols[3].selectbox("Soyad Sütunu", df.columns, 3),
        'advisor': cols[4].selectbox("Advisor Sütunu", df.columns, 4 if len(df.columns)>4 else 0)
    }
    
    classes = st.multiselect("İşlenecek Sınıfları Seçin", df[col_map['class']].unique())
    
    if classes:
        tmp_file = st.file_uploader("Master Şablon Dosyası", type=["xlsx"])
        
        if tmp_file and st.button("DOSYALARI OLUŞTUR", type="primary"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v9(t_bytes, c, sub_df, col_map, mod_in, debug_mode)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.balloons()
            st.success("✅ Tüm işlemler başarıyla tamamlandı!")
            st.download_button("📥 ZIP İndir", z_buf.getvalue(), "Gradebook_v9_Final.zip", "application/zip")
