import pandas as pd
from crudmodul.database import save_data

def delete_data(dfs):
    print("\n--- HAPUS DATA ---")
    print("1. Hapus ST (Otomatis menghapus Tim & Anggota terkait)")
    print("2. Hapus Tim (Otomatis menghapus Anggota terkait)")
    print("3. Hapus Anggota")
    choice = input("Pilih menu hapus (1-3): ")

    if choice == '1':
        print(dfs['ST'].to_string(index=False))
        try:
            idst = int(input("Masukkan idST yang akan dihapus: "))
        except ValueError:
            return
        dfs['ST'] = dfs['ST'][dfs['ST']['idST'] != idst]
        tims_to_del = dfs['tim'][dfs['tim']['idST'] == idst]['idtim']
        dfs['tim'] = dfs['tim'][dfs['tim']['idST'] != idst]
        dfs['anggota'] = dfs['anggota'][~dfs['anggota']['idtim'].isin(tims_to_del)]
        save_data(dfs)
        print("[SUKSES] ST beserta relasinya berhasil dihapus.")

    elif choice == '2':
        print(dfs['tim'].to_string(index=False))
        try:
            idtim = int(input("Masukkan idtim yang akan dihapus: "))
        except ValueError:
            return
        dfs['tim'] = dfs['tim'][dfs['tim']['idtim'] != idtim]
        dfs['anggota'] = dfs['anggota'][dfs['anggota']['idtim'] != idtim]
        save_data(dfs)
        print("[SUKSES] Tim beserta anggotanya berhasil dihapus.")

    elif choice == '3':
        print(dfs['anggota'].to_string(index=False))
        try:
            idang = int(input("Masukkan idang yang akan dihapus: "))
        except ValueError:
            return
        dfs['anggota'] = dfs['anggota'][dfs['anggota']['idang'] != idang]
        save_data(dfs)
        print("[SUKSES] Anggota berhasil dihapus.")
