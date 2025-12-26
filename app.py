import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Otomasyonu Pro", layout="wide")

# --- YARDIMCI FONKSİYONLAR ---

def copy_style(source_cell, target_cell):
    """Hücre stilini (Font, Kenarlık, Dolgu, Kilit, Hizalama) kopyalar."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def find_table_boundaries(ws):
    """
    Tablonun başını (Header) ve sonunu (Footer/Boşluk) bulur.
    Böylece aradaki boş 30 satırı tespit edip silebiliriz.
    """
    header_row = 0
    start_row = 5 # Varsayılan güvenlik
    end_row = ws.max_row
    
    # 1. Başlangıcı Bul: "Student Number" veya "Index" içeren satırı ara
    for row in ws.iter_rows(min_row=1, max_row=15):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if "Student" in cell.value or "Index" in cell.value or "Numara" in cell.value:
                    header_row = cell.row
                    start_row = header_row + 1 # Veri başlığın bir altından başlar
                    break
        if header_row > 0: break
    
    # 2. Bitişi Bul: start_row'dan aşağı inip tablonun nerede bittiğine bak.
    # Genelde "Total", "Average", "Advisor" yazar veya kenarlık biter.
    # Biz basitçe: İlk boş veya özel kelime içeren satırı bulalım.
    
    current_row = start_row
    max_search = 100 # Sonsuz döngü engeli
    
    while current_row < start_row + max_search:
        # Satırdaki A, B, C sütunlarına bak
        cell_a = ws.cell(row=current_row, column=1).value
        cell_b = ws.cell(row=current_row, column=2).value
        
        # Eğer hücrede "Advisor", "Total", "Average" varsa veya hücre tamamen boşsa ve border yoksa
        val_str = str(cell_a) if cell_a else ""
        if "Advisor" in val_str or "Total" in val_str or "Ortalama" in val_str:
            end_row = current_row
            break
        
        # Eğer şablonda 30 tane boş satır varsa, bunların hepsi boştur.
        # Ancak biz şablondaki o boşlukları silmek istiyoruz.
        # O yüzden manuel bir bitiş belirleyicisinden ziyade,
        # Kullanıcı şablonuna sadık kalarak, dolu olan son satırı bulup gerisini temizlemek daha güvenli.
        
        current_row += 1
        
    return start_row, end_row

def update_headers_and_names(wb, class_name, module_name, advisor_name):
    # Sheet ismini ve başlıkları güncelle (Önceki mantıkla aynı)
    main_ws = wb.worksheets[0]
    try:
        safe_title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        main_ws.title = safe_title
    except: pass

    for row in main_ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            val_str = str(cell.value)
            if "GRADEBOOK" in val_str and "MODULE" in val_str:
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            if "Advisor:" in val_str:
                cell.value = f"Advisor: {advisor_name}"

def process_class(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    # Advisor
    try: advisor_name = students_df.iloc[0][col_map['advisor']]
    except: advisor_name = "Belirtilmedi"

    update_headers_and_names(wb, class_name, module_name, advisor_name)

    # --- TABLO İŞLEME MANTIĞI ---
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 1. Tablonun sınırlarını belirle
        start_row, footer_row = find_table_boundaries(ws)
        
        # Şablondaki mevcut boş satır sayısı (Örn: 5. satırdan 35. satıra kadar boşsa 30 satır)
        # footer_row, "Advisor" yazan satır olsun. Veri alanı: start_row -> footer_row - 1
        
        # --- RESIZE STRATEJİSİ ---
        # En temiz yöntem: 
        # 1. İlk veri satırını (start_row) koru (Referans Satırı).
        # 2. Referans satırının ALTINDAKİ, footer'a kadar olan tüm boş satırları SİL.
        # 3. Öğrenci sayısı kadar yeni satır EKLE.
        
        rows_to_delete = footer_row - (start_row + 1)
        if rows_to_delete > 0:
            ws.delete_rows(start_row + 1, amount=rows_to_delete)
            
        # Şu an tablomuzda sadece 1 satır veri alanı kaldı (start_row).
        # Şimdi ihtiyacımız olan kadarını ekleyeceğiz.
        
        num_students = len(students_df)
        rows_to_add = num_students - 1 
        
        if rows_to_add > 0:
            # start_row'un altına ekle
            ws.insert_rows(start_row + 1, amount=rows_to_add)
            
        # --- VERİ VE FORMÜL DÖKÜMÜ ---
        max_col = ws.max_column
        
        for i, (_, student) in enumerate(students_df.iterrows()):
            current_row = start_row + i
            
            # Stil ve Formül Kopyalama (İlk satırdan diğerlerine)
            if i > 0:
                for col in range(1, max_col + 1):
                    source_cell = ws.cell(row=start_row, column=col) # Referans: İlk satır
                    target_cell = ws.cell(row=current_row, column=col) # Hedef: Yeni satır
                    
                    copy_style(source_cell, target_cell)
                    
                    # --- FORMÜL KAYDIRMA (TRANSLATOR) ---
                    if source_cell.data_type == 'f':
                        try:
                            # Formülü yeni konuma göre tercüme et (B3 -> B4)
                            target_cell.value = Translator(
                                source_cell.value, 
                                origin=source_cell.coordinate
                            ).translate_formula(target_cell.coordinate)
                        except:
                            # Çeviremezse olduğu gibi kopyala (fallback)
                            target_cell.value = source_cell.value

            # Öğrenci Bilgileri (Formül değilse yaz)
            # Not: Eğer şablonda B sütununda formül varsa üzerine yazmamalıyız.
            # Genelde No, Ad, Soyad sütunları boş olur, formül olmaz.
            
            ws.cell(row=current_row, column=1).value = i + 1
            ws.cell(row=current_row, column=2).value = student[col_map['no']]
            ws.cell(row=current_row, column=3).value = student[col_map['name']]
            ws.cell(row=current_row, column=4).value = student[col_map['surname']]

    # Dosya Kayıt İşlemleri (Aynı)
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    sheets_to_keep = ["MidTerm", "MET", "Midterm"]
    sheets_to_delete = [s for s in wb.sheetnames if s not in sheets_to_keep]
    for s in sheets_to_delete: del wb[s]
        
    checker_io = io.BytesIO()
    if len(wb.sheetnames) > 0:
        wb.save(checker_io)
        checker_io.seek(0)
    else:
        checker_io = None

    return main_io, checker_io

# --- ARAYÜZ ---
st.title("🎓 Otomatik Gradebook Pro v3.0")
st.markdown("**Yenilikler:** Akıllı Tablo Boyutlandırma + Formül Kaydırma")

tabs = st.tabs(["🚀 Oluştur", "ℹ️ Format"])

with tabs[0]:
    st.header("1. Ayarlar")
    module_input = st.text_input("Modül İsmi", "MODULE 2")
    
    st.header("2. Liste ve Şablon")
    student_file = st.file_uploader("Öğrenci Listesi", type=["xlsx"])

    if student_file:
        df = pd.read_excel(student_file)
        st.info("Sütun Eşleştirme:")
        cols = st.columns(5)
        class_col = cols[0].selectbox("Sınıf", df.columns, index=0)
        no_col = cols[1].selectbox("Numara", df.columns, index=1 if len(df.columns)>1 else 0)
        name_col = cols[2].selectbox("Ad", df.columns, index=2 if len(df.columns)>2 else 0)
        surname_col = cols[3].selectbox("Soyad", df.columns, index=3 if len(df.columns)>3 else 0)
        advisor_col = cols[4].selectbox("Advisor", df.columns, index=4 if len(df.columns)>4 else 0)
        
        col_mapping = {'class': class_col, 'no': no_col, 'name': name_col, 'surname': surname_col, 'advisor': advisor_col}

        selected_classes = st.multiselect("Sınıfları Seçin", df[class_col].unique())
        
        if selected_classes:
            st.warning("Master Şablon (Formüllü ve Boş 1 Satır Örnekli)")
            template_file = st.file_uploader("Şablon Yükle", type=["xlsx"])
            
            if template_file and st.button("Başlat", type="primary"):
                progress = st.progress(0)
                zip_buffer = io.BytesIO()
                template_bytes = template_file.getvalue()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for i, sinif in enumerate(selected_classes):
                        class_df = df[df[class_col] == sinif].reset_index(drop=True)
                        main, checker = process_class(template_bytes, sinif, class_df, col_mapping, module_input)
                        
                        zf.writestr(f"{sinif}/{sinif} GRADEBOOK.xlsx", main.getvalue())
                        if checker:
                            zf.writestr(f"{sinif}/{sinif} 1st Checker.xlsx", checker.getvalue())
                            zf.writestr(f"{sinif}/{sinif} 2nd Checker.xlsx", checker.getvalue())
                        
                        progress.progress((i + 1) / len(selected_classes))
                
                st.success("İşlem Tamam!")
                st.download_button("ZIP İndir", zip_buffer.getvalue(), "Gradebooks_Pro.zip", "application/zip")

with tabs[1]:
    st.markdown("""
    ### Önemli: Şablon Nasıl Olmalı?
    1. **Tek Satır Örnek:** Şablonunuzda öğrenci listesi için **en az 1 satır** (Örn: 5. Satır) ayrılmış olmalı.
    2. **Fazlalıklar:** Şablonunuzda 30 boş satır olsa bile program bunları **otomatik silip** sınıf mevcudu kadar (örn: 18) satır açacaktır.
    3. **Bitiş Sınırı:** Program tablonun bittiğini anlamak için "Advisor", "Total" gibi yazıları veya boş satırları arar.
    """)
