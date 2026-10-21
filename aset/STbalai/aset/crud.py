from crudmodul.database import load_data, clear_screen
from crudmodul.view import view_data
from crudmodul.create import add_st, add_tim, add_anggota
from crudmodul.delete import delete_data

def main():
    while True:
        clear_screen()
        print("=========================================")
        print("   APLIKASI CRUD RELATIONAL EXCEL (ST)   ")
        print("=========================================")
        print("1. Lihat Semua Data (Struktur Tree/Root)")
        print("2. Tambah Surat Tugas (ST)")
        print("3. Tambah Tim di dalam ST")
        print("4. Tambah Anggota di dalam Tim")
        print("5. Hapus Data (Delete)")
        print("6. Keluar")
        print("=========================================")
        choice = input("Pilih menu (1-6): ")

        dfs = load_data()

        if choice == '1':
            view_data(dfs)
        elif choice == '2':
            add_st(dfs)
            input("\nTekan Enter untuk melanjutkan...")
        elif choice == '3':
            add_tim(dfs)
            input("\nTekan Enter untuk melanjutkan...")
        elif choice == '4':
            add_anggota(dfs)
            input("\nTekan Enter untuk melanjutkan...")
        elif choice == '5':
            delete_data(dfs)
            input("\nTekan Enter untuk melanjutkan...")
        elif choice == '6':
            print("Keluar dari program. Terima kasih!")
            break
        else:
            input("Pilihan tidak valid. Tekan Enter untuk mengulang...")

if __name__ == '__main__':
    main()
