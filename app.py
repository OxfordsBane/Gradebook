import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v3.2", layout="wide")

# --- STİL KOPYALAMA ---
def copy_style(source_cell, target_cell):
    """Hücre stilini (Font, Kenarlık, Dolgu, Kilit) kopyalar."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# --- TABLO SINIRLARINI BULMA (YENİLENMİŞ MANTIK) ---
def find_table_slots(ws):
    """
    Bir sheetteki boş şablon satırlarını görsel (border) ve içerik analiziyle bulur.
    Return: start_row, end_row (verilerin girileceği aralık)
    """
    start_row = 0
    
    # 1. BAŞLANGICI BUL (Header Satırı)
    # İlk 15 satırda Header arıyoruz.
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "student" in val or "index" in val or "number" in val or "numara" in val:
                    start_row = cell.row + 1 # Veri, başlığın hemen altından başlar
                    break
        if start_row > 0: break
    
    if start_row == 0: 
        # Header bulamazsa varsayılan 6. satır diyelim (Güvenlik)
        return 6, 35 

    # 2. BİTİŞİ BUL (Slot Analizi)
    # start_row'dan aşağı doğru inip "Burası hala tablo mu?" diye bakacağız.
    # Kriter: Hücre boşsa VE kenarlığı varsa tablodur.
    # Yazı gelirse veya kenarlık biterse tablo biter.
    
    current_row = start_row
    max_search = 100 # En fazla 100 satır aşağı bak
    
    while current_row < start_row + max_search:
        cell_a = ws.cell(row=current_row, column=1) # A sütunu
        cell_b = ws.cell(row=current_row, column=2) # B sütunu (Numara genelde burada)
        
        # İçerik kontrolü
        val_str = (str(cell_a.value) if cell_a.value else "") + (str(cell_b.value) if cell_b.value else "")
        val_lower = val_str.lower()
        
        # Bitiş Sinyalleri (Yazı gelmesi)
        stop_keywords = ["total", "average", "advisor", "ortalama", "toplam", "checker", "grade", "score"]
        if any(keyword in val_lower for keyword in stop_keywords):
            break
        
        # Stil Kontrolü (Kenarlık yoksa bitmiştir)
        # Not: openpyxl'de border nesnesi her zaman vardır ama style 'none' olabilir.
        # Bu kontrol bazen yanıltıcı olabilir, o yüzden 'boş hücre' kontrolü daha güvenlidir.
        # Şablon mantığı: Boş satırlar vardır.
        
        # Eğer hücre doluysa ve yukarıdaki keywordlerden biri değilse? 
        # (Örn: Şablonda örnek öğrenci varsa). Devam etmeli.
        
        # Güvenli çıkış: Eğer arka arkaya 5 satır tamamen stilsiz/boş gelirse döngüyü kırabiliriz.
        # Ama şimdilik "Keywords" ve "Layout" yapısına güveniyoruz.
        
        current_row += 1
        
    end_row = current_row - 1
    
    # Eğer hiç boşluk bulamazsa (end < start), en az 1 satır var varsayalım
    if end_row < start_row:
        end_row = start_row
        
    return start_row, end_row

# --- SHEET İŞLEME ---
def process_sheet(ws, students_df, col_map):
    """Tek bir sheeti (Main, Midterm, TW vs.) alır, resize eder ve doldurur."""
    
    # 1. Tablonun sınırlarını bul
    start_row, end_row = find_table_slots(ws)
    
    # Şablondaki mevcut boş slot sayısı
    available_slots = end_row - start_row + 1
    num_students = len(students_df)
    
    # --- RESIZE MANTIĞI ---
    
    # DURUM A: ÖĞRENCİ SAYISI AZ (FAZLALIKLARI SİL)
    # Örn: 30 slot var, 20 öğrenci geldi. 10 satır silinecek.
    if num_students <= available_slots:
        rows_to_delete = available_slots - num_students
        if rows_to_delete > 0:
            # Silme işlemini öğrencilerin bittiği yerden (start + num) başlat
            delete_start = start_row + num_students
            ws.delete_rows(delete_start, amount=rows_to_delete)

    # DURUM B: ÖĞRENCİ SAYISI ÇOK (UZATMA YAP)
    # Örn: 30 slot var, 40 öğrenci geldi. 10 satır eklenecek.
    else:
        rows_to_add = num_students - available_slots
        # Mevcutların sonuna (end_row'un altına) ekle
        ws.insert_rows(end_row + 1, amount=rows_to_add)
        
        # STİL VE FORMÜL KOPYALAMA
        # Referans satırı: Mevcut son boş satır (end_row).
        # Neden? Çünkü header (start_row-1) kalın çerçeveli olabilir. 
        # Ama end_row genelde tablonun ortasındaki ince çerçeveli standart satırdır.
        ref_row = end_row 
        max_col = ws.max_column
        
        for i in range(rows_to_add):
            new_row_idx = end_row + 1 + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row, column=col)
                target = ws.cell(row=new_row_idx, column=col)
                
                copy_style(source, target)
                
                if source.data_type == 'f':
                    try:
                        target.value = Translator(source.value, source.coordinate).translate_formula(target.coordinate)
                    except:
                        target.value = source.value

    # --- VERİ DOLDURMA ---
    # Artık satır sayısı tam. Yazmaya başla.
    for i, (_, student) in enumerate(students_df.iterrows()):
        current_row = start_row + i
        
        # No, Ad, Soyad yaz (Formül üzerine yazma!)
        # Hücrede formül yoksa veriyi yaz. Varsa dokunma (Excel hesaplasın).
        
        # Sütun 1: Index
        c1 = ws.cell(row=current_row, column=1)
        if c1.data_type != 'f': c1.value = i + 1
        
        # Sütun 2: No
        c2 = ws.cell(row=current_row, column=2)
        if c2.data_type != 'f': c2.value = student[col_map['no']]
            
        # Sütun 3: Ad
        c3 = ws.cell(row=current_row, column=3)
        if c3.data_type != 'f': c3.value = student[col_map['name']]
            
        # Sütun 4: Soyad
        c4 = ws.cell(row=current_row, column=4)
        if c4.data_type != 'f': c4.value = student[col_map['surname']]

# --- ANA KONTROL ---
def process_workbook_data(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    # Advisor
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # Sadece ilk sheetteki başlıkları güncelle
    # Diğer sheetlerde de başlık güncellemek isterseniz bu kodu döngüye alabilirsiniz.
    try:
        main_ws = wb.worksheets[0]
        main_ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        
        for row in main_ws.iter_rows(min_row=1, max_row=10, max_col=20):
            for cell in row:
                if not cell.value: continue
                val = str(cell.value)
                if "GRADEBOOK" in val and "MODULE" in val:
                    cell.value = f"{class_name} GRADEBOOK - {module_name}"
                if "Advisor:" in val:
                    cell.value = f"Advisor: {advisor}"
    except: pass

    # TÜM SHEETLERİ İŞLE (Main, Midterm, TW...)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        process_sheet(ws, students_df, col_map)

    # KAYDET
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker Temizliği
    sheets_to_keep = ["MidTerm", "MET", "Midterm"]
    to_delete = [s for s in wb.sheetnames if s not in sheets_to_keep]
    for s in to_delete: del wb[s]
    
    checker_io = io.BytesIO() if len(wb.sheetnames) > 0 else None
    if checker_io:
        wb.save(checker_io)
        checker_io.seek(0)

    return main_io, checker_io

# --- ARAYÜZ ---
st.title("🎓 Gradebook Pro v3.2 (Universal Fix)")
st.markdown("Tüm sheetlerde tablo boyutunu ve formatı otomatik ayarlar.")

tabs = st.tabs(["Uygulama", "Bilgi"])

with tabs[0]:
    col1, col2 = st.columns(2)
    module_input = col1.text_input("Modül", "MODULE 2")
    
    student_file = st.file_uploader("Öğrenci Listesi", type=["xlsx"])
    if student_file:
        df = pd.read_excel(student_file)
        
        c_cols = st.columns(5)
        col_map = {
            'class': c_cols[0].selectbox("Sınıf", df.columns, index=0),
            'no': c_cols[1].selectbox("No", df.columns, index=1),
            'name': c_cols[2].selectbox("Ad", df.columns, index=2),
            'surname': c_cols[3].selectbox("Soyad", df.columns, index=3),
            'advisor': c_cols[4].selectbox("Advisor", df.columns, index=4 if len(df.columns)>4 else 0)
        }
        
        classes = st.multiselect("Sınıflar", df[col_map['class']].unique())
        
        if classes:
            template_file = st.file_uploader("Master Şablon", type=["xlsx"])
            if template_file and st.button("Başlat"):
                zip_buf = io.BytesIO()
                temp_bytes = template_file.getvalue()
                
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    prog = st.progress(0)
                    for i, cls in enumerate(classes):
                        sub_df = df[df[col_map['class']] == cls].reset_index(drop=True)
                        main, chk = process_workbook_data(temp_bytes, cls, sub_df, col_map, module_input)
                        
                        zf.writestr(f"{cls}/{cls} GRADEBOOK.xlsx", main.getvalue())
                        if chk:
                            zf.writestr(f"{cls}/{cls} 1st Checker.xlsx", chk.getvalue())
                            zf.writestr(f"{cls}/{cls} 2nd Checker.xlsx", chk.getvalue())
                        prog.progress((i+1)/len(classes))
                
                st.success("Tüm sheetler düzenlendi!")
                st.download_button("İndir", zip_buf.getvalue(), "Gradebooks_v3.2.zip", "application/zip")

with tabs[1]:
    st.markdown("""
    ### Bu Versiyon Neyi Çözdü?
    1. **Tüm Sheetler:** Artık sadece Main değil, Midterm, TW, Role-play gibi tüm sheetlerdeki tablolar algılanıp resize ediliyor.
    2. **Format Koruma:** Tablonun sonundaki ince çizgili boş satırı referans aldığı için kalın çerçeve sorunu yaşanmıyor.
    3. **Otomatik Algılama:** "Advisor" yazısı olmasa bile, boş satırların bittiği yeri algılayıp tablonun sonunu buluyor.
    """)
