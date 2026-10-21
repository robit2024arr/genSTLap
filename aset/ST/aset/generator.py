import os
import re
from docx import Document
from docx.oxml import parse_xml
import openpyxl

# Konfigurasi Nama File dan Folder
EXCEL_FILE = "aset/ST/data.xlsx"
TEMPLATE_DOCX = "aset/ST/ST-template.docx"
OUTPUT_DIR = "ST"
SHEET_ST = "ST"
SHEET_TIM = "tim"
SHEET_ANGGOTA = "anggota"

def clean_val(val):
    """Membersihkan nilai sel agar jika kosong/nan menjadi string kosong."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "nat", "none", "<na>"]:
        return ""
    return s

def sheet_to_dict_list(sheet):
    """Mengubah worksheet openpyxl menjadi list of dictionary berdasarkan header di baris pertama."""
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    result = []
    
    for row in rows[1:]:
        row_dict = {}
        for idx, header in enumerate(headers):
            if header:
                val = row[idx] if idx < len(row) else None
                row_dict[header] = clean_val(val)
        result.append(row_dict)
        
    return result

def replace_text_in_paragraph(paragraph, replacements):
    """Fungsi penggantian teks yang mempertahankan format font asli pada run serta mendukung list (multiparagraf)."""
    for key, val in replacements.items():
        if key in paragraph.text:
            if isinstance(val, list):
                if not val:
                    for r in paragraph.runs:
                        r.text = ""
                    continue
                
                # Simpan style dan posisi Ruler (Indent/Spacing) dari paragraf template
                orig_style = paragraph.style
                orig_left_indent = paragraph.paragraph_format.left_indent
                orig_first_line_indent = paragraph.paragraph_format.first_line_indent
                orig_space_after = paragraph.paragraph_format.space_after
                orig_line_spacing = paragraph.paragraph_format.line_spacing
                
                f_name, f_size, f_bold, f_italic, f_color = None, None, None, None, None
                if paragraph.runs:
                    orig_run = paragraph.runs[0]
                    f_name = orig_run.font.name
                    f_size = orig_run.font.size
                    f_bold = orig_run.font.bold
                    f_italic = orig_run.font.italic
                    if orig_run.font.color and orig_run.font.color.rgb:
                        f_color = orig_run.font.color.rgb

                # Kosongkan teks paragraf placeholder lama
                for r in paragraph.runs:
                    r.text = ""

                # Sisipkan baris berurutan dari atas ke bawah
                ref_paragraph = paragraph
                for line_text in val:
                    new_p = ref_paragraph.insert_paragraph_before(line_text)
                    new_p.style = orig_style
                    
                    # Terapkan posisi Ruler / Indentasi dari template
                    new_p.paragraph_format.left_indent = orig_left_indent
                    new_p.paragraph_format.first_line_indent = orig_first_line_indent
                    if orig_space_after is not None:
                        new_p.paragraph_format.space_after = orig_space_after
                    if orig_line_spacing is not None:
                        new_p.paragraph_format.line_spacing = orig_line_spacing
                    
                    # Terapkan format font pada run pertama
                    if new_p.runs:
                        run = new_p.runs[0]
                        run.font.name = f_name
                        run.font.size = f_size
                        run.font.bold = f_bold
                        run.font.italic = f_italic
                        if f_color:
                            run.font.color.rgb = f_color
                    
                    # Pindahkan acuan agar baris berikutnya disisipkan DI BAWAHNYA secara berurutan
                    if ref_paragraph != paragraph:
                        ref_paragraph._element.addnext(new_p._element)
                    ref_paragraph = new_p
                
                # Hapus paragraf bekas placeholder
                p_element = paragraph._element
                p_element.getparent().remove(p_element)
                break

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
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        
        if not all(sheet in wb.sheetnames for sheet in [SHEET_ST, SHEET_TIM, SHEET_ANGGOTA]):
            print(f"Error: Salah satu sheet ({SHEET_ST}, {SHEET_TIM}, {SHEET_ANGGOTA}) tidak ditemukan di file Excel.")
            return

        list_st = sheet_to_dict_list(wb[SHEET_ST])
        list_tim = sheet_to_dict_list(wb[SHEET_TIM])
        list_anggota = sheet_to_dict_list(wb[SHEET_ANGGOTA])
        
    except Exception as e:
        print(f"Gagal membaca file Excel: {e}")
        return

    for row_st in list_st:
        primary_id = clean_val(row_st.get('idST'))
        nost = clean_val(row_st.get('nost'))
        tgl = clean_val(row_st.get('Tgl'))
        
        raw_filename = f"{nost}-{tgl}"
        safe_filename = re.sub(r'[\\/*?:"<>|]', "_", raw_filename)
        file_name = f"{safe_filename}.docx"
        output_filename = os.path.join(OUTPUT_DIR, file_name)
        
        doc = Document(TEMPLATE_DOCX)
        
        tim_rows = [t for t in list_tim if t.get('idST') == primary_id]
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
                    for t_idx, tim_data in enumerate(tim_rows):
                        current_row = active_rows[t_idx]
                        
                        urut = urut + 1
                        idtim_val = clean_val(tim_data.get('idtim'))
                        anggota_list = [a for a in list_anggota if a.get('idtim') == idtim_val]
                        
                        anggota_list_text = []
                        str_posisi = ""
                        for a_idx, ang_row in enumerate(anggota_list):
                            anggota_list_text.append(f"{a_idx+1}.\t{clean_val(ang_row.get('nama'))}")
                            str_posisi += f"{clean_val(ang_row.get('posisi'))}\n"
                            
                        str_posisi = str_posisi.strip()
                        
                        row_replacements = {
                            "{{namatim}}": str(urut),
                            "{{jl}}": clean_val(tim_data.get('jl')),
                            "{{kel}}": clean_val(tim_data.get('kel')),
                            "{{kec}}": clean_val(tim_data.get('kec')),
                            "{{kab}}": clean_val(tim_data.get('kab')),
                            "{{s}}": f"S: {clean_val(tim_data.get('s'))}" if clean_val(tim_data.get('s')) else "",
                            "{{e}}": f"E: {clean_val(tim_data.get('e'))}" if clean_val(tim_data.get('e')) else "",
                            "{{s1}}": f"S: {clean_val(tim_data.get('s1'))}" if clean_val(tim_data.get('s1')) else "",
                            "{{e1}}": f"E: {clean_val(tim_data.get('e1'))}" if clean_val(tim_data.get('e1')) else "",
                            "{{pelaksanaan}}": clean_val(tim_data.get('pelaksanaan')),
                            "{{nama}}": anggota_list_text,
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
