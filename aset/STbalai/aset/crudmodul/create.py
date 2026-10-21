import pandas as pd
from crudmodul.database import save_data
from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter

def add_st(dfs):
    print("\n--- TAMBAH SURAT TUGAS (ST) ---")
    tgl = input("Masukkan Tanggal (Tgl): ")
    nost = input("Masukkan Nomor ST (nost): ")
    try:
        lamahari = int(input("Masukkan Lama Hari (lamahari): "))
    except ValueError:
        print("[ERROR] Lama hari harus berupa angka.")
        return

    st_df = dfs['ST']
    new_id = int(st_df['idST'].max()) + 1 if not st_df.empty and pd.notna(st_df['idST'].max()) else 1
    new_row = {'idST': new_id, 'Tgl': tgl, 'nost': nost, 'lamahari': lamahari}
    dfs['ST'] = pd.concat([st_df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(dfs)
    print(f"[SUKSES] ST baru ditambahkan dengan ID: {new_id}")

def add_tim(dfs):
    print("\n--- TAMBAH TIM DI DALAM ST ---")
    st_df = dfs['ST']
    if st_df.empty:
        print("[ERROR] Belum ada data ST. Harap tambah ST terlebih dahulu.")
        return
    
    tim_df = dfs['tim']
    
    print(f"{'ID ST':<8} | {'Nomor ST':<30} | {'Jumlah Tim':<12}")
    print("-" * 56)
    for _, st in st_df.iterrows():
        idst = st['idST']
        count_tim = len(tim_df[tim_df['idST'] == idst])
        print(f"{idst:<8} | {str(st['nost']):<30} | {count_tim:<12}")
    print("-" * 56)
    
    try:
        idst = int(input("\nPilih ID ST induk: "))
    except ValueError:
        print("[ERROR] ID ST harus berupa angka.")
        return
        
    if idst not in st_df['idST'].values:
        print("[ERROR] ID ST tidak ditemukan.")
        return

    existing_tim_in_st = tim_df[tim_df['idST'] == idst]
    next_tim_num = int(existing_tim_in_st['tim'].max()) + 1 if not existing_tim_in_st.empty and pd.notna(existing_tim_in_st['tim'].max()) else 1
    
    print(f"\n[INFO] Surat Tugas ini sudah memiliki {len(existing_tim_in_st)} Tim.")
    print(f"[INFO] Tim baru akan otomatis diset sebagai: Tim Ke-{next_tim_num}")
    
    jl = input("Masukkan Jalan (jl): ")
    kel = input("Masukkan Kelurahan (kel): ")

    new_idtim = int(tim_df['idtim'].max()) + 1 if not tim_df.empty and pd.notna(tim_df['idtim'].max()) else 1

    new_row = {
        'idtim': new_idtim,
        'idST': idst,
        'tim': next_tim_num,
        'jl': jl,
        'kel': kel,
        'kec': '-',
        'kab': '-',
        's': 0,
        'e': 0,
        'pelaksanaan': '-'
    }
    dfs['tim'] = pd.concat([tim_df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(dfs)
    print(f"[SUKSES] Tim baru (Tim Ke-{next_tim_num}) berhasil ditambahkan ke ST ID {idst}!")

def add_anggota(dfs):
    print("\n--- TAMBAH ANGGOTA DI DALAM TIM ---")
    st_df = dfs['ST']
    tim_df = dfs['tim']
    anggota_df = dfs['anggota']

    if tim_df.empty:
        print("[ERROR] Belum ada data Tim. Harap tambah Tim terlebih dahulu.")
        return

    print("\nPilih Tim Berdasarkan Struktur Pohon di Bawah Ini:")
    print("=" * 80)
    
    for idx, st in st_df.iterrows():
        idst = st['idST']
        print(f"📄 [ST ID: {idst}] No. ST: {st['nost']} | Tanggal ST: {st['Tgl']}")
        
        tims = tim_df[tim_df['idST'] == idst]
        if tims.empty:
            print("    └── ⚠️ Belum ada Tim")
        else:
            for t_idx, tim in tims.iterrows():
                idtim = tim['idtim']
                is_last_tim = (t_idx == tims.index[-1])
                tim_prefix = "└──" if is_last_tim else "├──"
                
                count_ang = len(anggota_df[anggota_df['idtim'] == idtim])
                print(f"    {tim_prefix} 👥 [ID Tim: {idtim}] Tim Ke-{tim['tim']} | Jl. {tim['jl']}, Kel. {tim['kel']} | 👤 Jumlah Anggota: {count_ang}")
    
    print("=" * 80)

    try:
        idtim = int(input("\nPilih ID Tim induk tempat anggota akan ditambahkan: "))
    except ValueError:
        print("[ERROR] ID Tim harus berupa angka.")
        return
        
    if idtim not in tim_df['idtim'].values:
        print("[ERROR] ID Tim tidak ditemukan.")
        return

    # Memuat data autocomplete dari file autocomplete.xlsx, sheet 'anggota', kolom B (indeks 1) mulai baris ke-2
    try:
        auto_df = pd.read_excel('autocomplete.xlsx', sheet_name='anggota')
        autocomplete_names = auto_df.iloc[:, 1].dropna().astype(str).tolist()
    except Exception:
        autocomplete_names = []

    # Membuat WordCompleter dari prompt_toolkit agar daftar pilihan muncul otomatis saat mengetik
    anggota_completer = WordCompleter(autocomplete_names, ignore_case=True)

    # Perulangan untuk menambah anggota berturut-turut pada tim yang sama
    while True:
        # Menggunakan prompt() dari prompt_toolkit untuk fitur as-you-type autocomplete
        nama = prompt("Masukkan Nama Anggota: ", completer=anggota_completer)
        
        posisi = input("Masukkan Posisi (contoh: ketua / anggota) [Default: anggota]: ").strip()
        if not posisi:
            posisi = 'anggota'

        # Refresh referensi anggota_df terbaru setiap iterasi penambahan
        anggota_df = dfs['anggota']
        new_idang = int(anggota_df['idang'].max()) + 1 if not anggota_df.empty and pd.notna(anggota_df['idang'].max()) else 1
        
        tim_row = tim_df[tim_df['idtim'] == idtim]
        namatim_val = int(tim_row['tim'].values[0]) if not tim_row.empty else 1

        new_row = {
            'idang': new_idang,
            'idtim': idtim,
            'namatim': namatim_val,
            'nama': nama,
            'posisi': posisi
        }
        dfs['anggota'] = pd.concat([anggota_df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(dfs)
        print(f"[SUKSES] Anggota baru berhasil ditambahkan dengan ID: {new_idang}")

        # Menampilkan daftar anggota yang sudah ada di tim ini saat ini
        print("\n" + "-" * 50)
        print(f"📋 Daftar Anggota Saat Ini di Tim ID {idtim}:")
        current_members = dfs['anggota'][dfs['anggota']['idtim'] == idtim]
        if current_members.empty:
            print("   (Belum ada anggota)")
        else:
            for idx_m, m in enumerate(current_members.itertuples(), 1):
                print(f"   {idx_m}. {m.nama} ({m.posisi})")
        print("-" * 50)

        # Pertanyaan untuk melanjutkan penambahan anggota di tim yang sama
        lanjut = input("Ingin menambah nama lain di tim ini? (y/n): ").strip().lower()
        if lanjut != 'y':
            break
