import pandas as pd
import os
from crudmodul.database import clear_screen

def view_data(dfs):
    clear_screen()
    print("=" * 80)
    print("                 STRUKTUR POHON DATA (SURAT TUGAS ➔ TIM ➔ ANGGOTA)")
    print("=" * 80)
    
    st_df = dfs['ST']
    tim_df = dfs['tim']
    anggota_df = dfs['anggota']

    if st_df.empty:
        print("\n[INFO] Belum ada data Surat Tugas.")
        input("\nTekan Enter untuk kembali ke menu...")
        return

    for idx, st in st_df.iterrows():
        idst = st['idST']
        print(f"\n📄 [ST ID: {idst}] No. ST: {st['nost']}")
        print(f"   ├── Tanggal    : {st['Tgl']}")
        print(f"   └── Durasi     : {st['lamahari']} Hari")
        
        tims = tim_df[tim_df['idST'] == idst]
        if tims.empty:
            print("        └── ⚠️ Belum ada Tim terdaftar")
        else:
            for t_idx, tim in tims.iterrows():
                idtim = tim['idtim']
                is_last_tim = (t_idx == tims.index[-1])
                tim_prefix = "└──" if is_last_tim else "├──"
                
                print(f"        {tim_prefix} 👥 [TIM ID: {idtim}] Tim Ke-{tim['tim']}")
                print(f"            ├── Lokasi : Jl. {tim['jl']}, Kel. {tim['kel']}")
                
                anggotas = anggota_df[anggota_df['idtim'] == idtim]
                if anggotas.empty:
                    print(f"            └── ⚠️ Belum ada Anggota")
                else:
                    for a_idx, ang in anggotas.iterrows():
                        is_last_ang = (a_idx == anggotas.index[-1])
                        ang_prefix = "└──" if is_last_ang else "├──"
                        print(f"            {ang_prefix} 👤 [Anggota ID: {ang['idang']}] {ang['nama']} ({ang['posisi']})")

    print("\n" + "=" * 80)
    input("Tekan Enter untuk kembali ke menu...")
