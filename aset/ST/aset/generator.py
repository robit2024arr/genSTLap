import os
import re
import pandas as pd
from docx import Document
from docx.oxml import parse_xml

# Konfigurasi Nama File dan Folder
EXCEL_FILE = "aset/ST/data.xlsx"
TEMPLATE_DOCX = "aset/ST/ST-template.docx"
OUTPUT_DIR = "ST"
SHEET_ST = "ST"
SHEET_TIM = "tim"
SHEET_ANGGOTA = "anggota"

def clean_val(val):
    """Membersihkan nilai sel agar jika kosong/nan menjadi string kosong."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "nat", "none", "<na>"]:
        return ""
    return s

def replace_text_in_paragraph(paragraph, replacements):
    """Fungsi penggantian teks yang mempertahankan format font asli pada run."""
    for key, val in replacements.items():
        if key in paragraph.text:
            replaced_in_run = False
            for run in paragraph.runs:
                if key in run.text:
                    f_name = run.font.name
                    f_size = run.font.size
                    f_bold = run.font.bold
                    f_italic = run.font.italic
                    f_color = run.font.color.rgb if run.font.color else None
                    
                    run.text = run.text.replace(key, val)
                    
                    run.font.name = f_name
                    run.font.size = f_size
                    run.font.bold = f_bold
                    run.font.italic = f_italic
                    if f_color:
                        run.font.color.rgb = f_color
                        
                    replaced_in_run = True
                    break
            
            if not replaced_in_run:
                full_text = "".join([run.text for run in paragraph.runs])
                if key in full_text:
                    full_text = full_text.replace(key, val)
                    for i, run in enumerate(paragraph.runs):
                        if i == 0:
                            f_name = run.font.name
                            f_size = run.font.size
                            f_bold = run.font.bold
                            f_italic = run.font.italic
                            
                            run.text = full_text
                            
                            run.font.name = f_name
                            run.font.size = f_size
                            run.font.bold = f_bold
                            run.font.italic = f_italic
                        else:
                            run.text = ""

def main():
    # Cek apakah file Excel dan Template ada
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: File Excel '{EXCEL_FILE}' tidak ditemukan!")
        return

    if not os.path.exists(TEMPLATE_DOCX):
        print(f"Error: Template Word '{TEMPLATE_DOCX}' tidak ditemukan!")
        return

    # Buat folder output jika belum ada
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("Memulai proses generasi dokumen dengan format nama file 'nost-Tgl.docx'...")

    try:
        # Membaca seluruh data sebagai string (dtype=str) agar format tampilan asli 
        # (seperti koordinat minus, titik, koma) terjaga, dan fillna("") untuk sel kosong.
        df_st = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_ST, header=0, dtype=str).fillna("")
        df_tim = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_TIM, header=0, dtype=str).fillna("")
        df_anggota = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_ANGGOTA, header=0, dtype=str).fillna("")
    except Exception as e:
        print(f"Gagal membaca file Excel: {e}")
        return

    for index, row_st in df_st.iterrows():
        primary_id = clean_val(row_st.get('idST'))
        nost = clean_val(row_st.get('nost'))
        tgl = clean_val(row_st.get('Tgl'))
        
        # Membuat format nama file: nost-Tgl.docx (membersihkan karakter ilegal sistem file)
        raw_filename = f"{nost}-{tgl}"
        safe_filename = re.sub(r'[\\/*?:"<>|]', "_", raw_filename)
        file_name = f"{safe_filename}.docx"
        output_filename = os.path.join(OUTPUT_DIR, file_name)
        
        doc = Document(TEMPLATE_DOCX)
        
        tim_rows = df_tim[df_tim['idST'] == primary_id].reset_index(drop=True)
        num_tim = len(tim_rows)
        
        replacements = {
            "{{nost}}": nost,
            "{{Tgl}}": tgl,
            "{{lamahari}}": clean_val(row_st.get('lamahari')),
        }

        for p in doc.paragraphs:
            replace_text_in_paragraph(p, replacements)
            
        for section in doc.sections:
            if section.header:
                for p in section.header.paragraphs:
                    replace_text_in_paragraph(p, replacements)
            if section.footer:
                for p in section.footer.paragraphs:
                    replace_text_in_paragraph(p, replacements)

        for table in doc.tables:
            is_tim_table = False
            row_template_idx = -1
            for r_idx, row_t in enumerate(table.rows):
                row_text = "".join([cell.text for cell in row_t.cells])
                if "{{namatim}}" in row_text or "{{jl}}" in row_text:
                    is_tim_table = True
                    row_template_idx = r_idx
                    break
            
            if is_tim_table and row_template_idx != -1:
                template_row = table.rows[row_template_idx]
                
                if num_tim == 0:
                    for cell in template_row.cells:
                        cell.text = ""
                else:
                    cell_xmls = [cell._tc.xml for cell in template_row.cells]
                    
                    active_rows = []
                    for t_idx in range(num_tim):
                        if t_idx == 0:
                            active_rows.append(template_row)
                        else:
                            new_row = table.add_row()
                            for c_idx, xml_str in enumerate(cell_xmls):
                                new_cell = new_row.cells[c_idx]
                                new_tc = parse_xml(xml_str)
                                new_cell._tc.getparent().replace(new_cell._tc, new_tc)
                            active_rows.append(new_row)
                    
                    urut = 0
                    for t_idx, tim_data in tim_rows.iterrows():
                        current_row = active_rows[t_idx]
                        
                        urut = urut + 1
                        idtim_val = clean_val(tim_data['idtim'])
                        anggota_list = df_anggota[df_anggota['idtim'] == idtim_val].reset_index(drop=True)
                        
                        str_anggota = ""
                        str_posisi = ""
                        for a_idx, ang_row in anggota_list.iterrows():
                            str_anggota += f"{a_idx+1}. {clean_val(ang_row['nama'])}\n"
                            str_posisi += f"{clean_val(ang_row['posisi'])}\n"
                            
                        str_anggota = str_anggota.strip()
                        str_posisi = str_posisi.strip()
                        
                        row_replacements = {
                            "{{namatim}}": str(urut),
                            "{{jl}}": clean_val(tim_data.get('jl')),
                            "{{kel}}": clean_val(tim_data.get('kel')),
                            "{{kec}}": clean_val(tim_data.get('kec')),
                            "{{kab}}": clean_val(tim_data.get('kab')),
                            # Menambahkan label S: dan E: secara otomatis
                            "{{s}}": f"S: {clean_val(tim_data.get('s'))}" if clean_val(tim_data.get('s')) else "",
                            "{{e}}": f"E: {clean_val(tim_data.get('e'))}" if clean_val(tim_data.get('e')) else "",
                            "{{s1}}": f"S: {clean_val(tim_data.get('s1'))}" if clean_val(tim_data.get('s1')) else "",
                            "{{e1}}": f"E: {clean_val(tim_data.get('e1'))}" if clean_val(tim_data.get('e1')) else "",
                            "{{pelaksanaan}}": clean_val(tim_data.get('pelaksanaan')),
                            "{{nama}}": str_anggota,
                            "{{posisi}}": str_posisi
                        }
                        
                        for cell in current_row.cells:
                            for p in cell.paragraphs:
                                replace_text_in_paragraph(p, row_replacements)
            else:
                for row_t in table.rows:
                    for cell in row_t.cells:
                        for p in cell.paragraphs:
                            replace_text_in_paragraph(p, replacements)

        doc.save(output_filename)
        print(f"[BERHASIL] Dibuat: {output_filename}")

    print(f"Semua dokumen berhasil digenerate dengan format nama 'nost-Tgl.docx' di dalam folder '{OUTPUT_DIR}'!")

if __name__ == "__main__":
    main()
