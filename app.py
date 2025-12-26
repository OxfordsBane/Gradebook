import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v3.1", layout="wide")

# --- STİL KOPYALAMA ---
def copy_style(source_cell, target_cell):
    """Hücre stilini birebir kopyalar."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# --- TABLO SINIRLARINI BULMA ---
def find_available_rows(ws):
    """
    Şablondaki boş veri satırlarının başlangıcını ve bitişini bulur.
    Örn: 5. satırdan başlayıp 35. satıra kadar boş hücreler varsa bunları tespit eder.
    """
    start_row = 0
    end_row = 0
    
    # 1. Başlangıcı Bul (Header'dan sonraki ilk satır)
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if "Index" in cell.value or "Student" in cell.value or "No" in str(cell.value):
                    start_row = cell.row + 1
                    break
        if start_row > 0: break
    
    if start_row == 0: start_row = 5 # Bulamazsa varsayılan
    
    # 2. Bitişi Bul (Advisor/Total yazısına kadar olan boşluk)
    # start_row'dan aşağı iniyoruz.
    current = start_row
    max_look = 200
    
    while current < start_row + max_look:
        # A ve B sütununu kontrol et
        val_a = ws.cell(row=current, column=1).value
        val_b = ws.cell(row=current, column=2).value
        val_str = str(val_a) if val_a else "" + str(val_b) if val_b else ""
        
        # Bitiş sinyalleri
        if "Advisor" in val_str or "Total" in val_str or "Ortalama" in val_str:
            end_row = current - 1
            break
        
        # Eğer satırın alt kenarlığı kalınsa bu da bir bitiş işaretidir (Opsiyonel)
        # Şimdilik sadece metin tabanlı bitiş yapıyoruz.
        
        current += 1
        
    if end_row == 0: end_row = start_row + 29 # Bulamazsa 30 satır varsay
    
    return start_row, end_row

# --- BAŞLIKLARI GÜNCELLEME ---
def update_headers(ws, class_name, module_name, advisor_name):
    try:
        ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
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
def process_class(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # Sadece ilk sheetteki başlıkları güncelle (Genelde main sheet)
    update_headers(wb.worksheets[0], class_name, module_name, advisor)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 1. Mevcut Boşlukları Tespit Et
        start_row, end_row = find_available_rows(ws)
        available_slots = end_row - start_row + 1
        num_students = len(students_df)
        
        # --- DURUM 1: ÖĞRENCİ SAYISI AZ (FAZLALIKLARI SİL) ---
        if num_students <= available_slots:
            # Önce öğrencileri mevcut satırlara yaz
            limit_row = start_row + num_students
            
            # Geriye kalan boş satırları sil (Tabloyu yukarı çek)
            rows_to_delete = available_slots - num_students
            if rows_to_delete > 0:
                # Silme işlemini öğrencilerin bittiği yerin altından yap
                ws.delete_rows(limit_row, amount=rows_to_delete)

        # --- DURUM 2: ÖĞRENCİ SAYISI ÇOK (UZATMA YAP) ---
        else:
            rows_to_add = num_students - available_slots
            # Mevcutların sonuna ekleme yap
            ws.insert_rows(end_row + 1, amount=rows_to_add)
            
            # Yeni eklenen satırlara STİL KOPYALA
            # Stil kaynağı olarak "end_row"u (mevcut son boş satırı) kullanıyoruz.
            # Bu satır genelde "orta" stilindedir (ince kenarlık), header değildir.
            ref_row = end_row
            max_col = ws.max_column
            
            for i in range(rows_to_add):
                new_row_idx = end_row + 1 + i
                for col in range(1, max_col + 1):
                    source = ws.cell(row=ref_row, column=col)
                    target = ws.cell(row=new_row_idx, column=col)
                    
                    copy_style(source, target)
                    
                    # Formül Kaydırma
                    if source.data_type == 'f':
                        try:
                            target.value = Translator(
                                source.value, source.coordinate
                            ).translate_formula(target.coordinate)
                        except:
                            target.value = source.value

        # --- VERİLERİ YAZMA DÖNGÜSÜ ---
        # Artık satır sayısı tam ayarlandı, sırayla yazabiliriz.
        for i, (_, student) in enumerate(students_df.iterrows()):
            current_row = start_row + i
            
            # No, Ad, Soyad yaz
            ws.cell(row=current_row, column=1).value = i + 1
            ws.cell(row=current_row, column=2).value = student[col_map['no']]
            ws.cell(row=current_row, column=3).value = student[col_map['name']]
            ws.cell(row=current_row, column=4).value = student[col_map['surname']]

    # KAYDETME İŞLEMLERİ
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker temizliği
    for s in [s for s in wb.sheetnames if s not in ["MidTerm", "MET", "Midterm"]]:
        del wb[s]
    
    checker_io = io.BytesIO() if len(wb.sheetnames) > 0 else None
    if checker_io:
        wb.save(checker_io)
        checker_io.seek(0)

    return main_io, checker_io

# --- ARAYÜZ ---
st.title("🎓 Gradebook Pro v3.1 (Smart Fill)")
st.markdown("Format bozulmadan mevcut satırları doldurur, fazlalığı siler veya uzatır.")

tabs = st.tabs(["İşlem", "Nasıl Çalışır?"])

with tabs[0]:
    col_set1, col_set2 = st.columns(2)
    module_input = col_set1.text_input("Modül", "MODULE 2")
    
    student_file = st.file_uploader("Öğrenci Listesi", type=["xlsx"])
    if student_file:
        df = pd.read_excel(student_file)
        
        c1, c2, c3, c4, c5 = st.columns(5)
        col_map = {
            'class': c1.selectbox("Sınıf", df.columns, index=0),
            'no': c2.selectbox("No", df.columns, index=1),
            'name': c3.selectbox("Ad", df.columns, index=2),
            'surname': c4.selectbox("Soyad", df.columns, index=3),
            'advisor': c5.selectbox("Advisor", df.columns, index=4 if len(df.columns)>4 else 0)
        }
        
        classes = st.multiselect("Sınıflar", df[col_map['class']].unique())
        
        if classes:
            template_file = st.file_uploader("Şablon (30 satırlık boş hali)", type=["xlsx"])
            if template_file and st.button("Başlat"):
                zip_buf = io.BytesIO()
                temp_bytes = template_file.getvalue()
                
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    prog = st.progress(0)
                    for i, cls in enumerate(classes):
                        sub_df = df[df[col_map['class']] == cls].reset_index(drop=True)
                        main, chk = process_class(temp_bytes, cls, sub_df, col_map, module_input)
                        
                        zf.writestr(f"{cls}/{cls} GRADEBOOK.xlsx", main.getvalue())
                        if chk:
                            zf.writestr(f"{cls}/{cls} 1st Checker.xlsx", chk.getvalue())
                            zf.writestr(f"{cls}/{cls} 2nd Checker.xlsx", chk.getvalue())
                        prog.progress((i+1)/len(classes))
                
                st.download_button("ZIP İndir", zip_buf.getvalue(), "Gradebooks.zip", "application/zip")

with tabs[1]:
    st.markdown("""
    **Format Koruma Mantığı:**
    Bu versiyon şablonu silip baştan yapmaz.
    1. Şablonunuzdaki 30 (veya kaç taneyse) boş satırı bulur.
    2. Öğrencileri bu satırlara yazar.
    3. Eğer öğrenci sayısı azsa (örn: 20), kalan 10 boş satırı siler.
       *Böylece en üstteki ve en alttaki özel çizgiler bozulmaz.*
    4. Eğer öğrenci sayısı fazlaysa (örn: 35), sona 5 satır ekler ve stili **son satırdan** kopyalar.
    """)
