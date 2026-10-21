import os
import re
import openpyxl
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, Cm

# Nama file Excel, Template Word, dan Folder Output
EXCEL_FILE = "aset/Lap/data.xlsx"
TEMPLATE_DOCX = "aset/Lap/laporan-template.docx"
OUTPUT_DIR = "Lap"
SHEET_ST = "ST"
SHEET_TIM = "tim"
SHEET_ANGGOTA = "anggota"

# Cek apakah file Excel dan Template ada
if not os.path.exists(EXCEL_FILE):
    print(f"Error: File Excel '{EXCEL_FILE}' tidak ditemukan!")
    exit(1)

if not os.path.exists(TEMPLATE_DOCX):
    print(f"Error: Template Word '{TEMPLATE_DOCX}' tidak ditemukan!")
    exit(1)

# Buat folder output jika belum ada
os.makedirs(OUTPUT_DIR, exist_ok=True)

print("Memulai proses generasi dokumen...")

# Fungsi helper untuk membaca worksheet menjadi list of dictionaries
def load_sheet_as_dict(workbook, sheet_name):
    if sheet_name not in workbook.sheetnames:
        return [], []
    sheet = workbook[sheet_name]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], []
    
    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    data = []
    for row in rows[1:]:
        row_dict = {}
        for col_idx, cell_value in enumerate(row):
            header = headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
            # Obati None menjadi string kosong atau nilai teks
            if cell_value is None:
                row_dict[header] = ""
            else:
                row_dict[header] = str(cell_value).strip()
        data.append(row_dict)
    return data, headers

try:
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    list_st, headers_st = load_sheet_as_dict(wb, SHEET_ST)
    list_tim, headers_tim = load_sheet_as_dict(wb, SHEET_TIM)
    list_anggota, headers_anggota = load_sheet_as_dict(wb, SHEET_ANGGOTA)
except Exception as e:
    print(f"Gagal membaca file Excel: {e}")
    exit(1)

def get_run_font_style(run):
    """Mengambil properti styling font dari sebuah run."""
    if not run:
        return None
    return {
        'name': run.font.name,
        'size': run.font.size,
        'bold': run.font.bold,
        'italic': run.font.italic,
        'underline': run.font.underline,
        'color': run.font.color.rgb if run.font.color else None
    }

def apply_font_style(run, style):
    """Menerapkan properti styling font ke sebuah run."""
    if not style or not run:
        return
    if style['name']: run.font.name = style['name']
    if style['size']: run.font.size = style['size']
    if style['bold'] is not None: run.font.bold = style['bold']
    if style['italic'] is not None: run.font.italic = style['italic']
    if style['underline'] is not None: run.font.underline = style['underline']
    if style['color']: run.font.color.rgb = style['color']

def replace_text_withbullet(paragraph, key, value, delimiter=","):
    """Mengganti placeholder dengan daftar bullet sambil mempertahankan gaya font asli."""
    if key in paragraph.text:
        items = [item.strip() for item in str(value).split(delimiter) if item.strip()]
        
        # Ambil sampel gaya font dari run yang mengandung key
        sample_font = None
        for run in paragraph.runs:
            if key in run.text:
                sample_font = get_run_font_style(run)
                break

        if not items:
            paragraph.text = paragraph.text.replace(key, "")
            return

        paragraph.text = paragraph.text.replace(key, "")
        p_element = paragraph._p
        current_p = paragraph

        for item in items:
            bullet_item = f"{item}"
            new_p = current_p.insert_paragraph_before(bullet_item)
            new_p.paragraph_format.left_indent = Cm(1.5)
            
            if sample_font:
                for run in new_p.runs:
                    apply_font_style(run, sample_font)

            parent = p_element.getparent()
            parent.remove(new_p._p)
            p_element.addnext(new_p._p)
            p_element = new_p._p

def replace_text_in_paragraph(paragraph, replacements):
    """Mengganti placeholder di paragraf dengan mempertahankan format asli run."""
    for key, val in replacements.items():
        if key in paragraph.text:
            replaced_in_run = False
            for run in paragraph.runs:
                if key in run.text:
                    f_style = get_run_font_style(run)
                    run.text = run.text.replace(key, val)
                    apply_font_style(run, f_style)
                    replaced_in_run = True
                    break
            
            if not replaced_in_run:
                full_text = "".join([run.text for run in paragraph.runs])
                if key in full_text:
                    full_text = full_text.replace(key, val)
                    for i, run in enumerate(paragraph.runs):
                        if i == 0:
                            f_style = get_run_font_style(run)
                            run.text = full_text
                            apply_font_style(run, f_style)
                        else:
                            run.text = ""

def set_cell_text_with_style(cell, text, style):
    """Mengisi teks ke sel tabel dengan mempertahankan gaya font dari template sel."""
    for p in cell.paragraphs:
        p.text = ""
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
    
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(text)
    apply_font_style(run, style)

def set_cell_vmerge(cell, val):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('vMerge'):
            tcPr.remove(child)
    if val == 'restart':
        vmerge = parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>')
        tcPr.append(vmerge)
    elif val == 'continue':
        vmerge = parse_xml(f'<w:vMerge {nsdecls("w")}/>')
        tcPr.append(vmerge)

# Iterasi setiap baris dari sheet ST
for row_st in list_st:
    primary_id = row_st.get('idST', '')
    idst_str = str(primary_id).strip()
    nost = str(row_st.get('nost', '')).strip()
    kel = str(row_st.get('kel', '')).strip()
    tgl_st = str(row_st.get('Tgl', '')).strip()
    
    # Format nama file output baru
    raw_filename = f"{idst_str}.pem.{kel} {tgl_st}"
    safe_filename = re.sub(r'[\\/*?:"<>|]', "_", raw_filename)
    file_name = f"{safe_filename}.docx"
    output_filename = os.path.join(OUTPUT_DIR, file_name)
    
    doc = Document(TEMPLATE_DOCX)
    
    # Filter tim_rows berdasarkan primary_id (idST)
    tim_rows = [row for row in list_tim if str(row.get('idST', '')).strip() == idst_str]
    num_tim = len(tim_rows)
    
    replacements = {
        "[[nost]]": nost,
        "[[Tgl]]": tgl_st,
        "[[Tglkapital]]": tgl_st.upper(),
        "[[lamahari]]": str(row_st.get('lamahari', '')),
        "[[jl]]": str(row_st.get('jl', '')),
        "[[kel]]": str(row_st.get('kel', '')),
        "[[kec]]": str(row_st.get('kec', '')),
        "[[kab]]": str(row_st.get('kab', '')),
        "[[ketua]]": str(row_st.get('ketua', '')),
        "[[nipketua]]": str(row_st.get('nipketua', '')),
        "[[s]]": str(row_st.get('s', '')),
        "[[e]]": str(row_st.get('e', '')),
        "[[cuaca]]": str(row_st.get('cuaca', '')),
        "[[suhu]]": str(row_st.get('suhu', '')),
        "[[kelembaban]]": str(row_st.get('kelembaban', '')),
        "[[curahhujan]]": str(row_st.get('curahhujan', '')),
        "[[angin]]": str(row_st.get('angin', '')),
        "[[batasadministrasi]]": str(row_st.get('batasadministrasi', '')),
        "[[laporandari]]": str(row_st.get('laporandari', '')),
        "[[kawasan]]": str(row_st.get('kawasan', '')),
        "[[mulai]]": str(row_st.get('mulai', '')),
        "[[totalterbakar]]": str(row_st.get('totalterbakar', '')),
        "[[totaldipadamkan]]": str(row_st.get('totaldipadamkan', '')),
        "[[sumberair]]": str(row_st.get('sumberair', '')),
        "[[tipekebakaran]]": str(row_st.get('tipekebakaran', '')),
        "[[jenistanah]]": str(row_st.get('jenistanah', '')),
        "[[vegetasi]]": str(row_st.get('vegetasi', '')),
        "[[SelesaiPemadaman]]": str(row_st.get('SelesaiPemadaman', '')),
        "[[kondisiapi]]": str(row_st.get('kondisiapi', '')),
        "[[kembalike]]": str(row_st.get('kembalike', '')),
        "[[Kendala]]": str(row_st.get('Kendala', '')),
        "[[pekerjaanmasyarakat]]": str(row_st.get('pekerjaanmasyarakat', '')),
        "[[Aktivitasmasyarkat]]": str(row_st.get('Aktivitasmasyarkat', '')),
        "[[AsalTim]]": str(row_st.get('AsalTim', '')),
    }

    for p in doc.paragraphs:
        replace_text_withbullet(p, "[[alat]]", row_st.get("alat", ""))
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
            if "[[namatim]]" in row_text or "[[jl]]" in row_text or "[[nama]]" in row_text or "[[nip]]" in row_text:
                is_tim_table = True
                row_template_idx = r_idx
                break
        
        if is_tim_table and row_template_idx != -1:
            template_row = table.rows[row_template_idx]
            
            # Rekam gaya font asli dari setiap kolom template baris tabel
            column_styles = {}
            for c_idx, cell in enumerate(template_row.cells):
                if cell.paragraphs and cell.paragraphs[0].runs:
                    column_styles[c_idx] = get_run_font_style(cell.paragraphs[0].runs[0])
                else:
                    column_styles[c_idx] = None

            if num_tim == 0:
                for cell in template_row.cells:
                    cell.text = ""
            else:
                nama_col_indices = []
                posisi_col_indices = []
                no_col_indices = []
                tanggal_col_indices = []
                lokasi_col_indices = []
                ket_col_indices = []
                
                for c_idx, cell in enumerate(template_row.cells):
                    cell_text = cell.text.lower()
                    if "[[nama]]" in cell.text:
                        nama_col_indices.append(c_idx)
                    elif "[[nip]]" in cell.text:
                        posisi_col_indices.append(c_idx)
                    elif "no" in cell_text or c_idx == 0:
                        no_col_indices.append(c_idx)
                    elif "tanggal" in cell_text or "tgl" in cell_text:
                        tanggal_col_indices.append(c_idx)
                    elif "keterangan" in cell_text or "ket" in cell_text or c_idx == len(template_row.cells) - 1:
                        ket_col_indices.append(c_idx)
                    else:
                        lokasi_col_indices.append(c_idx)

                shared_col_indices = no_col_indices + tanggal_col_indices + lokasi_col_indices + ket_col_indices

                team_blocks = []
                total_physical_rows = 0

                for tim_data in tim_rows:
                    # Ambil nama kolom B dari sheet tim secara dinamis (indeks ke-1)
                    kolom_b_tim_name = headers_tim[1] if len(headers_tim) > 1 else ""
                    id_tim_val = str(tim_data.get(kolom_b_tim_name, '')).strip()

                    # Ambil nama kolom B dari sheet anggota secara dinamis (indeks ke-1)
                    kolom_b_anggota_name = headers_anggota[1] if len(headers_anggota) > 1 else ""

                    # Filter sheet "anggota" di mana Kolom B-nya bernilai sama dengan Kolom B sheet "tim"
                    anggota_list = [
                        row for row in list_anggota
                        if str(row.get(kolom_b_anggota_name, '')).strip() == id_tim_val
                    ]
                    
                    n_ang = max(1, len(anggota_list))
                    
                    team_blocks.append({
                        'tim_data': tim_data,
                        'anggota_list': anggota_list,
                        'count': n_ang,
                        'start_row_idx': total_physical_rows
                    })
                    total_physical_rows += n_ang

                while len(table.rows) - row_template_idx < total_physical_rows:
                    table.add_row()
                urutrow = 0

                for block in team_blocks:
                    urutrow += 1
                    tim_data = block['tim_data']
                    anggota_list = block['anggota_list']
                    n_ang = block['count']
                    start_r = row_template_idx + block['start_row_idx']

                    tgl_pelaksanaan = str(tim_data.get('pelaksanaan', '')).strip()
                    if not tgl_pelaksanaan or tgl_pelaksanaan.lower() == 'nan' or tgl_pelaksanaan == '-':
                        tgl_pelaksanaan = tgl_st

                    jl = str(tim_data.get('jl', ''))
                    kel = str(tim_data.get('kel', ''))
                    kec = str(tim_data.get('kec', ''))
                    kab = str(tim_data.get('kab', ''))
                    s_val = str(tim_data.get('s', ''))
                    e_val = str(tim_data.get('e', ''))

                    loc_lines = []
                    if jl and jl != 'nan' and jl != '-': loc_lines.append(jl)
                    if kel and kel != 'nan' and kel != '-': loc_lines.append(kel)
                    if kec and kec != 'nan' and kec != '-': loc_lines.append(kec)
                    if kab and kab != 'nan' and kab != '-': loc_lines.append(kab)
                    loc_lines.append("Koordinat:")
                    loc_lines.append(f"S {s_val}")
                    loc_lines.append(f"E {e_val}")
                    loc_text_str = "\n".join(loc_lines)

                    for sub_i in range(n_ang):
                        current_row_idx = start_r + sub_i
                        current_row = table.rows[current_row_idx]

                        current_row.height = Pt(15)
                        trPr = current_row._tr.get_or_add_trPr()
                        trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="260" w:hRule="atLeast"/>')
                        trPr.append(trHeight)

                        for c_idx, cell in enumerate(current_row.cells):
                            tcPr = cell._tc.get_or_add_tcPr()
                            tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/></w:tcMar>')
                            tcPr.append(tcMar)

                        # Mengisi Kolom No
                        for c_idx in no_col_indices:
                            set_cell_text_with_style(current_row.cells[c_idx], str(urutrow), column_styles.get(c_idx))

                        # Mengisi Kolom Tanggal Pelaksanaan
                        for c_idx in tanggal_col_indices:
                            set_cell_text_with_style(current_row.cells[c_idx], tgl_pelaksanaan, column_styles.get(c_idx))

                        # Mengisi Kolom Lokasi Kegiatan & Koordinat
                        for c_idx in lokasi_col_indices:
                            set_cell_text_with_style(current_row.cells[c_idx], loc_text_str, column_styles.get(c_idx))

                        # Kosongkan kolom Keterangan
                        for c_idx in ket_col_indices:
                            set_cell_text_with_style(current_row.cells[c_idx], "", column_styles.get(c_idx))

                        # Mengisi kolom Nama dan Posisi Anggota
                        if len(anggota_list) > 0 and sub_i < len(anggota_list):
                            ang_row = anggota_list[sub_i]
                            ang_name_str = f"{sub_i+1}. {str(ang_row.get('nama', ''))}"
                            ang_pos_str = f"{str(ang_row.get('nip', ''))}"

                            for c_idx in nama_col_indices:
                                set_cell_text_with_style(current_row.cells[c_idx], ang_name_str, column_styles.get(c_idx))

                            for c_idx in posisi_col_indices:
                                set_cell_text_with_style(current_row.cells[c_idx], ang_pos_str, column_styles.get(c_idx))
                        else:
                            for c_idx in nama_col_indices:
                                current_row.cells[c_idx].text = ""
                            for c_idx in posisi_col_indices:
                                current_row.cells[c_idx].text = ""

                    if n_ang > 1:
                        for c_idx in shared_col_indices:
                            set_cell_vmerge(table.rows[start_r].cells[c_idx], 'restart')
                            for r_sub in range(1, n_ang):
                                set_cell_vmerge(table.rows[start_r + r_sub].cells[c_idx], 'continue')
        else:
            for row_t in table.rows:
                for cell in row_t.cells:
                    for p in cell.paragraphs:
                        replace_text_in_paragraph(p, replacements)

    doc.save(output_filename)
    print(f"[BERHASIL] Dibuat: {output_filename}")

print("Semua dokumen berhasil digenerate!")
