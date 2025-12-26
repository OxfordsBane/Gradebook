import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
from copy import copy
import io
import zipfile

st.set_page_config(page_title="Gradebook Pro v8.0 (Heartbeat Fix)", layout="wide")

# --- HÜCRE KLONLAMA ---
def clone_cell(source_cell, target_cell):
    """Stil ve Formül Kopyalar."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
    
    if source_cell.data_type == 'f':
        try:
            target_cell.value = Translator(
                source_cell.value, source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
        except:
            target_cell.value = source_cell.value
    elif source_cell.value is not None:
        # Formül değilse ve boş değilse (örn: "-" işareti) kopyala
         target_cell.value = source_cell.value

# --- TABLO BAŞLANGICINI BUL ---
def find_header_row(ws):
    """Sadece tablonun başladığı yeri bulur. Gerisi sabittir."""
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "index" in val or "student" in val or "number" in val or "no" in val:
                    return cell.row
    return 6 # Bulamazsa varsayılan

# --- GÜVENLİ RESIZE ---
def process_sheet_resize(ws, num_students):
    header_row = find_header_row(ws)
    
    # GÜVENLİ BÖLGE: Header'ın 5 satır altı.
    # Neden? Çünkü hemen altına eklersek bazen header'ın kalın çizgisini alabilir.
    # 5 satır altı (örn: 11. satır) kesinlikle tablonun "göbeğidir" ve standart formattadır.
    # Sizin "25. satıra ekliyorum" mantığınızla aynıdır, sadece biraz daha yukarıdadır.
    
    insert_pos = header_row + 5
    
    # Şablondaki mevcut boş satırları saymaya gerek yok mu?
    # VAR. Ama Footer'ı bulmak riskli olduğu için şöyle yapıyoruz:
    # Şablonun standart 30 satır olduğunu varsayıyoruz (veya kullanıcıdan alabiliriz).
    # Daha güvenli yol: Dolu satır sayısını kontrol et.
    
    # Basit ve Sağlam Yöntem:
    # Şablondaki mevcut satır sayısı (Veri alanı)
    # Bunu anlamak için insert_pos'tan aşağı doğru "Advisor" yazana kadar sayabiliriz.
    # Ama Advisor yazısı yoksa? 
    # Şöyle yapalım: Şablonda varsayılan olarak 30 boş satır olduğunu kabul edelim.
    # Bu genelde standarttır.
    
    current_capacity = 30 
    
    # Ancak kapasiteyi dinamik bulmak istersek:
    # insert_pos'tan aşağı 100 satır bak, kenarlık yoksa bitmiştir.
    check_row = insert_pos
    dynamic_cap = 0
    while check_row < insert_pos + 100:
        cell = ws.cell(row=check_row, column=1) # A sütunu
        # Eğer kenarlık varsa veya doluysa devam et
        if cell.border and (cell.border.left.style or cell.border.bottom.style or cell.value):
             dynamic_cap += 1
        else:
            # Kenarlık bittiyse tablo bitmiştir
            break
        check_row += 1
    
    # Eğer dinamik bulduysak onu kullan, yoksa 30 varsay
    if dynamic_cap > 5: 
        current_capacity = dynamic_cap + 5 # +5 çünkü yukarıdan başladık
    
    needed_rows = num_students
    
    # --- DURUM A: EKLEME ---
    if needed_rows > current_capacity:
        rows_to_add = needed_rows - current_capacity
        
        # Göbekten (insert_pos) ekleme yap
        ws.insert_rows(insert_pos, amount=rows_to_add)
        
        # Referans: Ekleme yerinin hemen üstü
        ref_row_idx = insert_pos - 1
        
        max_col = ws.max_column
        for i in range(rows_to_add):
            target_row_idx = insert_pos + i
            for col in range(1, max_col + 1):
                source = ws.cell(row=ref_row_idx, column=col)
                target = ws.cell(row=target_row_idx, column=col)
                clone_cell(source, target)

    # --- DURUM B: SİLME ---
    elif needed_rows < current_capacity:
        rows_to_delete = current_capacity - needed_rows
        # Silmeye yine güvenli bölgeden (insert_pos) başla
        # Bu sayede footer'a dokunmadan aradan çekmiş oluruz.
        ws.delete_rows(insert_pos, amount=rows_to_delete)
        
    # Veri giriş başlangıcı her zaman Header + 1'dir
    return header_row + 1

# --- BAŞLIK ---
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

# --- MAIN PROCESS ---
def process_workbook_v8(template_bytes, class_name, students_df, col_map, module_name):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    try: advisor = students_df.iloc[0][col_map['advisor']]
    except: advisor = ""

    # TÜM SHEETLER
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        update_headers(ws, class_name, module_name, advisor)
        
        # Resize yap
        data_start = process_sheet_resize(ws, len(students_df))
        
        # SADECE MAIN SHEET VERİ GİRİŞİ
        if wb.index(ws) == 0:
            for i, (_, student) in enumerate(students_df.iterrows()):
                r = data_start + i
                # Formülsüz hücrelere yaz
                if ws.cell(r, 1).data_type != 'f': ws.cell(r, 1).value = i + 1
                if ws.cell(r, 2).data_type != 'f': ws.cell(r, 2).value = student[col_map['no']]
                if ws.cell(r, 3).data_type != 'f': ws.cell(r, 3).value = student[col_map['name']]
                if ws.cell(r, 4).data_type != 'f': ws.cell(r, 4).value = student[col_map['surname']]

    # KAYDET
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
st.title("🎓 Gradebook Pro v8.0 (Heartbeat Insertion)")
st.markdown("Tablonun ortasından (güvenli bölgeden) ekleme yaparak footer'ı korur.")

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
    
    classes = st.multiselect("Sınıflar", df[col_map['class']].unique())
    if classes:
        tmp_file = st.file_uploader("Şablon", type=["xlsx"])
        if tmp_file and st.button("Başlat"):
            z_buf = io.BytesIO()
            t_bytes = tmp_file.getvalue()
            
            with zipfile.ZipFile(z_buf, "w") as zf:
                bar = st.progress(0)
                for i, c in enumerate(classes):
                    sub_df = df[df[col_map['class']] == c].reset_index(drop=True)
                    m, ch = process_workbook_v8(t_bytes, c, sub_df, col_map, mod_in)
                    
                    zf.writestr(f"{c}/{c} GRADEBOOK.xlsx", m.getvalue())
                    if ch:
                        zf.writestr(f"{c}/{c} 1st Checker.xlsx", ch.getvalue())
                        zf.writestr(f"{c}/{c} 2nd Checker.xlsx", ch.getvalue())
                    bar.progress((i+1)/len(classes))
            
            st.success("Tamamlandı!")
            st.download_button("İndir", z_buf.getvalue(), "Gradebook_v8.zip", "application/zip")
