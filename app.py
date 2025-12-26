import streamlit as st
import pandas as pd
import openpyxl
from copy import copy
import io
import zipfile

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gradebook Otomasyonu", layout="wide")

# --- YARDIMCI FONKSİYONLAR ---
def copy_style(source_cell, target_cell):
    """Hücre stilini (Font, Kenarlık, Dolgu, Kilit) kopyalar."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def update_headers_and_names(wb, class_name, module_name, advisor_name):
    """
    Şablondaki başlıkları, sheet ismini ve Advisor bilgisini günceller.
    """
    # 1. EN SOLDAKİ SHEET İSMİNİ DEĞİŞTİRME (MAIN SHEET)
    main_ws = wb.worksheets[0]
    try:
        # Excel sheet isimlerinde yasaklı karakterleri temizle (: \ / ? * [ ])
        safe_title = "".join([c for c in class_name if c not in r"[]:*?\/"])
        main_ws.title = safe_title
    except Exception as e:
        print(f"Sheet ismi değiştirilemedi: {e}")

    # 2. HÜCRE İÇERİKLERİNİ GÜNCELLEME (Smart Search)
    # İlk 10 satır ve ilk 20 sütunu tarayıp anahtar kelimeleri ararız.
    # Bu sayede şablonun formatı değişse bile kod çalışır.
    
    for row in main_ws.iter_rows(min_row=1, max_row=10, max_col=20):
        for cell in row:
            if not cell.value: continue
            
            val_str = str(cell.value)
            
            # Başlık Değişimi: "A1.02 GRADEBOOK - MODULE 2" formatını yakala
            if "GRADEBOOK" in val_str and "MODULE" in val_str:
                # Format: [SINIF] GRADEBOOK - [MODÜL]
                cell.value = f"{class_name} GRADEBOOK - {module_name}"
            
            # Advisor Değişimi: "Advisor:" içeren hücreyi yakala
            if "Advisor:" in val_str:
                cell.value = f"Advisor: {advisor_name}"

def process_class(template_bytes, class_name, students_df, col_map, module_name):
    """
    Hafızadaki şablonu işler ve çıktıları üretir.
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    
    # --- ADVISOR BİLGİSİNİ AL ---
    # Sınıf listesinde Advisor sütunundaki ilk değeri alıyoruz (Hepsi aynı varsayılır)
    try:
        advisor_name = students_df.iloc[0][col_map['advisor']]
    except:
        advisor_name = "Belirtilmedi"

    # --- BAŞLIKLARI VE İSİMLERİ GÜNCELLE ---
    update_headers_and_names(wb, class_name, module_name, advisor_name)

    # --- ÖĞRENCİLERİ EKLEME ---
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Referans satırı (Varsayım: 5. satırda veri başlıyor)
        start_row = 5 
        
        # Satır sayısını bulmak için basit bir kontrol (Opsiyonel: Dinamik yapılabilir)
        # Şimdilik sabit 5. satır varsayımıyla devam ediyoruz.

        num_students = len(students_df)
        rows_to_add = num_students - 1 
        
        if rows_to_add > 0:
            ws.insert_rows(start_row + 1, amount=rows_to_add)
            
        max_col = ws.max_column
        
        for i, (_, student) in enumerate(students_df.iterrows()):
            current_row = start_row + i
            
            # Stil Kopyalama
            if i > 0:
                for col in range(1, max_col + 1):
                    source_cell = ws.cell(row=start_row, column=col)
                    target_cell = ws.cell(row=current_row, column=col)
                    copy_style(source_cell, target_cell)
                    
                    if source_cell.data_type == 'f':
                        target_cell.value = source_cell.value 

            # Veri Yazma
            ws.cell(row=current_row, column=1).value = i + 1  # Index
            ws.cell(row=current_row, column=2).value = student[col_map['no']]
            ws.cell(row=current_row, column=3).value = student[col_map['name']]
            ws.cell(row=current_row, column=4).value = student[col_map['surname']]

    # Main Gradebook Kaydet
    main_io = io.BytesIO()
    wb.save(main_io)
    main_io.seek(0)
    
    # Checker Dosyaları (MidTerm ve MET dışındakileri sil)
    sheets_to_keep = ["MidTerm", "MET", "Midterm"]
    sheets_to_delete = [s for s in wb.sheetnames if s not in sheets_to_keep]
    
    for s in sheets_to_delete:
        del wb[s]
        
    checker_io = io.BytesIO()
    if len(wb.sheetnames) > 0:
        wb.save(checker_io)
        checker_io.seek(0)
    else:
        checker_io = None

    return main_io, checker_io

# --- ARAYÜZ (UI) ---

st.title("🎓 Otomatik Gradebook v2.0")
st.markdown("Sınıf isimlerini, modül bilgisini ve advisor ismini otomatik güncelleyen sürüm.")

tabs = st.tabs(["🚀 Gradebook Oluştur", "ℹ️ Bilgi"])

with tabs[0]:
    # --- ADIM 1: GENEL AYARLAR ---
    st.header("1. Genel Ayarlar")
    module_input = st.text_input("Şu anki Modül İsmi (Örn: MODULE 2)", "MODULE 2")

    st.divider()
    
    # --- ADIM 2: ÖĞRENCİ LİSTESİ ---
    st.header("2. Öğrenci Listesi Yükle")
    student_file = st.file_uploader("Tüm Sınıfların Listesi (Excel)", type=["xlsx", "xls"])

    if student_file:
        df = pd.read_excel(student_file)
        st.dataframe(df.head(3))
        
        st.info("Sütunları Eşleştirin:")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            class_col = st.selectbox("Sınıf", df.columns, index=0)
        with col2:
            no_col = st.selectbox("Numara", df.columns, index=1 if len(df.columns)>1 else 0)
        with col3:
            name_col = st.selectbox("Ad", df.columns, index=2 if len(df.columns)>2 else 0)
        with col4:
            surname_col = st.selectbox("Soyad", df.columns, index=3 if len(df.columns)>3 else 0)
        with col5:
            # Advisor seçimi - Eğer listede yoksa kullanıcı uyarılır
            advisor_col = st.selectbox("Advisor (Hoca)", df.columns, index=4 if len(df.columns)>4 else 0)
            
        col_mapping = {
            'class': class_col, 'no': no_col, 
            'name': name_col, 'surname': surname_col,
            'advisor': advisor_col
        }

        # --- ADIM 3: İŞLEM ---
        st.divider()
        st.header("3. Sınıf ve Şablon")
        
        unique_classes = df[class_col].unique()
        selected_classes = st.multiselect("İşlenecek Sınıfları Seçin", unique_classes)
        
        if selected_classes:
            st.warning(f"Seçili sınıflar için uygun MASTER ŞABLONU yükleyin.")
            template_file = st.file_uploader("Şablon Dosyası (.xlsx)", type=["xlsx"])
            
            if template_file and st.button("Dosyaları Oluştur", type="primary"):
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                zip_buffer = io.BytesIO()
                
                template_bytes = template_file.getvalue()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    total_classes = len(selected_classes)
                    
                    for idx, sinif in enumerate(selected_classes):
                        status_text.text(f"İşleniyor: {sinif}...")
                        
                        class_df = df[df[class_col] == sinif].reset_index(drop=True)
                        
                        # Process fonksiyonuna module_name'i de gönderiyoruz
                        main_io, checker_io = process_class(
                            template_bytes, sinif, class_df, col_mapping, module_input
                        )
                        
                        # Dosyaları ZIP'e ekle
                        zf.writestr(f"{sinif}/{sinif} GRADEBOOK.xlsx", main_io.getvalue())
                        
                        if checker_io:
                            zf.writestr(f"{sinif}/{sinif} 1st Checker Add-up.xlsx", checker_io.getvalue())
                            zf.writestr(f"{sinif}/{sinif} 2nd Checker Add-up.xlsx", checker_io.getvalue())
                        
                        progress_bar.progress((idx + 1) / total_classes)
                
                status_text.success("✅ Tüm işlemler tamamlandı!")
                st.download_button(
                    "📥 ZIP İndir", 
                    zip_buffer.getvalue(), 
                    "Gradebooks_v2.zip", 
                    "application/zip"
                )

with tabs[1]:
    st.markdown("""
    ### Yeni Özellikler Nasıl Çalışır?
    
    **1. Başlık Değişimi:**
    * Program, şablonun ilk sayfasında içinde **"GRADEBOOK"** ve **"MODULE"** kelimeleri geçen hücreyi arar.
    * Bulduğunda içeriği `[SINIF ADI] GRADEBOOK - [GİRDİĞİNİZ MODÜL]` olarak değiştirir.
    * *Örnek:* "A1.02 GRADEBOOK - MODULE 2"
    
    **2. Advisor (Hoca) İsmi:**
    * Yüklediğiniz sınıf listesinde hocaların isminin olduğu bir sütun olmalıdır.
    * Program şablonda **"Advisor:"** yazan hücreyi arar ve hocanın ismini oraya yazar.
    * *Örnek:* "Advisor: Kasım Burak Çavuşoğlu"
    
    **3. Sheet İsmi:**
    * Excel'i açtığınızda en altta görünen sekme ismi (Sheet Name), otomatik olarak sınıf kodu (Örn: A1.02) yapılır.
    """)
