import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v4.0 (Final Logic)", layout="wide")

# --- STİL VE FORMÜL KOPYALAMA ---
def copy_row_style_and_formula(ws, source_row_idx, target_row_idx):
    """
    Kaynak satırın (source_row) stilini ve formüllerini hedef satıra (target_row) kopyalar.
    Excel'deki 'Satırı Aşağı Sürükle' işleminin Python karşılığıdır.
    """
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        source_cell = ws.cell(row=source_row_idx, column=col)
        target_cell = ws.cell(row=target_row_idx, column=col)
        
        # 1. Stili Kopyala
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)
        
        # 2. Formül veya Değeri Kopyala
        if source_cell.data_type == 'f':
            # Formülse: Referansları kaydır (Örn: A5 -> A6)
            try:
                target_cell.value = Translator(
                    source_cell.value, source_cell.coordinate
                ).translate_formula(target_cell.coordinate)
            except:
                target_cell.value = source_cell.value # Çeviremezse aynısını yaz
        else:
            # Formül değilse: Sabit değerleri kopyalama (İsimler main sheette yazılacak)
            # Sadece Main Sheet dışındaki sayfalarda sabit metin varsa kopyalanabilir
            pass

# --- TABLO ALANINI BULMA ---
def find_template_range(ws):
    """
    Şablondaki veri girilecek alanı bulur.
    Start: Header'ın altı.
    End: Footer'ın (Total/Advisor) hemen üstü.
    """
    start_row = 0
    # 1. Başlangıcı Bul
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "student" in val or "index" in val or "number" in val:
                    start_row = cell.row + 1
                    break
        if start_row > 0: break
    
    if start_row == 0: start_row = 6 # Fallback
    
    # 2. Bitişi Bul (Advisor/Total yazısı veya boşluk bitimi)
    current = start_row
    max_look = 150
    end_row = start_row + 1
    
    found_footer = False
    while current < start_row + max_look:
        # A ve B sütununa bak
        val_a = str(ws.cell(row=current, column=1).value or "")
        val_b = str(ws.cell(row=current, column=2).value or "")
        val_combined = (val_a + val_b).lower()
        
        keywords = ["total", "advisor", "average", "toplam", "ortalama", "checker"]
        if any(k in val_combined for k in keywords):
            end_row = current - 1
            found_footer = True
            break
        current += 1
        
    if not found_footer:
        # Footer bulamazsa, stilin bittiği yeri tahmin etmeye çalışırız
        end_row = start_row + 29 # Varsayılan 30 satır
        
    return start_row, end_row

# --- BAŞLIKLARI GÜNCELLEME ---
def update_headers(ws, class_name, module_name, advisor_name):
    try:
        # Main sheet ismini sınıf adı yap
        if ws.parent.index(ws) == 0:
            ws.title = "".join([c for c in class_name if c not in r"[]:*?\/"])
    except: pass

    # Smart Search: Başlık ve Advisor
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            val = str(cell.value)
            if "GRADEBOOK" in val and "MODULE" in val:
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            if "Advisor:" in val:
                cell.value = f"Advisor: {advisor_name}"

# --- SAYFAYI YENİDEN BOYUTLANDIRMA (RESIZE) ---
def resize_sheet(ws, num_students):
    """
    Şablondaki satır sayısını öğrenci sayısına eşitler.
    Bunu yaparken "Insert Row" kullanır ve formülleri kopyalar.
    """
    start_row, end_row = find_template_range(ws)
    current_capacity = end_row - start_row + 1
    
    # Hedef satır sayısı
    needed_rows = num_students
    
    # DURUM 1: Kapasite Yetersiz -> Satır Ekle (INSERT)
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # Nereye ekleyeceğiz? Footer'ın hemen üstüne değil,
        # mevcut son satırın BİR ÜSTÜNE ekleyelim ki stil referansımız olsun.
        # En güvenlisi: start_row + 1 konumuna eklemek değil,
        # Listenin sonuna (end_row'a) ekleyip üstten kopyalamaktır.
        
        insert_pos = end_row 
        # Excel mantığı: Insert row dediğimizde o satır ve altındakiler aşağı kayar.
        
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Şimdi eklenen satırlara (insert_pos'tan insert_pos + rows_to_add'e kadar)
        # bir üst satırın (insert_pos - 1) özelliklerini kopyalayalım.
        source_row = insert_pos - 1
        
        for i in range(rows_to_add):
            target_row = insert_pos + i
            copy_row_style_and_formula(ws, source_row, target_row)
            
    # DURUM 2: Kapasite Fazla -> Satır Sil (DELETE)
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye sondan başla (Footer bozulmasın diye yukarıdan değil aşağıdan kırp)
        # Veri alanı: start_row ... end_row
        # Silinecek başlangıç: start_row + needed_rows
        
        delete_pos = start_row + needed_rows
        ws.delete_rows(delete_pos, amount=rows_to_delete)

    # İşlem sonrası yeni veri aralığı başlangıcı
    return start_row

# --- ANA İŞLEM ---
def process_workbook_logic(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # 1. TÜM SHEETLERİ GEZ VE BOYUTLANDIR
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Başlık güncelle
        update_headers(ws, class_name, module_name, advisor)
        
        # Resize İşlemi (Satır Ekle/Sil + Formül Taşı)
        data_start_row = resize_sheet(ws, len(students_df))
        
        # 2. VERİ GİRİŞİ (SADECE MAIN SHEET)
        # Diğer sheetler veriyi formülle çekeceği için onlara isim yazmıyoruz.
        if wb.index(ws) == 0: # Sadece ilk/ana sayfa
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start_row + i
                
                # Main Sheet'e verileri Hard-code olarak yazıyoruz
                # Formül varsa ezmemeye çalış, ama main sheette genelde isimler manuel girilir.
                
                # Index
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                # No
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                # Ad
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                # Soyad
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    # KAYDET
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker
    sheets_to_keep = ["MidTerm", "MET", "Midterm"]
    to_del = [s for s in wb.sheetnames if s not in sheets_to_keep]
    for s in to_del: del wb[s]
    
    chk_io = None
    if len(wb.sheetnames) > 0:
        chk_io = io.BytesIO()
        wb.save(chk_io)
        chk_io.seek(0)
        
    return main_io, chk_io

# --- ARAYÜZ ---
st.title("🎓 Gradebook Pro v4.0 (Manuel Yöntem Taklidi)")
st.markdown("""
**Çalışma Mantığı:**
1. Şablondaki satır sayısını kontrol eder.
2. Öğrenci sayısına göre **araya satır ekler** veya fazlalığı siler.
3. Eklenen satırlara **üst satırdaki formülleri** kopyalar.
4. İsimleri **sadece Ana Sayfaya** yazar (Diğer sayfalar formülle güncellenir).
""")

tabs = st.tabs(["Uygulama", "Önemli Notlar"])

with tabs[0]:
    c1, c2 = st.columns(2)
    module_input = c1.text_input("Modül", "MODULE 2")
    
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
            temp_file = st.file_uploader("Master Şablon", type=["xlsx"])
            if temp_file and st.button("Başlat"):
                z_buf = io.BytesIO()
                t_bytes = temp_file.getvalue()
                
                with zipfile.ZipFile(z_buf, "w") as zf:
                    bar = st.progress(0)
                    for i, cls in enumerate(classes):
                        sub_df = df[df[col_map['class']] == cls].reset_index(drop=True)
                        main, chk = process_workbook_logic(t_bytes, cls, sub_df, col_map, module_input)
                        
                        zf.writestr(f"{cls}/{cls} GRADEBOOK.xlsx", main.getvalue())
                        if chk:
                            zf.writestr(f"{cls}/{cls} 1st Checker.xlsx", chk.getvalue())
                            zf.writestr(f"{cls}/{cls} 2nd Checker.xlsx", chk.getvalue())
                        bar.progress((i+1)/len(classes))
                
                st.success("İşlem Tamam!")
                st.download_button("ZIP İndir", z_buf.getvalue(), "Gradebook_v4.zip", "application/zip")

with tabs[1]:
    st.warning("""
    **Şablon Hazırlığı İçin Kritik Bilgi:**
    
    Bu programın düzgün çalışması için, diğer sheetlerdeki (Midterm, TW vb.) öğrenci isim sütunlarının **FORMÜL İLE** Main Sheet'e bağlı olması gerekir.
    
    *Örn: TW1 sayfasındaki Ad hücresinde `='Main'!C6` gibi bir formül olmalıdır.*
    
    Program satır eklediğinde bu formülü aşağı çekecektir (Copy-Down). Eğer şablonunuzda formül yoksa, diğer sayfalarda isimler BOŞ çıkar.
    """)
