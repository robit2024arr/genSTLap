#!/bin/bash

while true; do
    clear
    echo "===================================================="
    echo "       SISTEM MANAJEMEN LAPORAN & ST - BAPPERIDA     "
    echo "===================================================="
    echo " [1] Sinkronisasi Data dari Spreadsheet Online"
    echo " [2] Jalankan Generator Laporan (Lap)"
    echo " [3] Jalankan Generator Surat Tugas (ST)"
    echo " [4] Menu Tambahan / Fitur Berikutnya (Placeholder)"
    echo "----------------------------------------------------"
    echo " [0] Keluar / Exit"
    echo "===================================================="
    
    read -p "Pilih menu [0-4]: " pilihan

    case $pilihan in
        1)
            echo ""
            echo "Menjalankan sinkronisasi data dari spreadsheet..."
            python3 aset/sh/ambildarispreadsheet.py
            echo ""
            echo "Tekan [Enter] untuk kembali ke menu..."
            read
            ;;
        2)
            echo ""
            echo "Menjalankan generator Laporan..."
            if [ -f "aset/Lap/genlap.py" ]; then
                python3 aset/Lap/genlap.py
            else
                echo "File aset/Lap/genlap.sh tidak ditemukan!"
            fi
            echo ""
            echo "Tekan [Enter] untuk kembali ke menu..."
            read
            ;;
        3)
            echo ""
            echo "Menjalankan generator Surat Tugas (ST)..."
            if [ -f "aset/ST/aset/generator.py" ]; then
                python3 aset/ST/aset/generator.py
            else
                echo "File aset/ST/generator.sh tidak ditemukan!"
            fi
            echo ""
            echo "Tekan [Enter] untuk kembali ke menu..."
            read
            ;;
        4)
            echo ""
            echo "Fitur berikutnya belum ditambahkan."
            echo ""
            echo "Tekan [Enter] untuk kembali ke menu..."
            read
            ;;
        0)
            echo ""
            echo "Keluar dari program. Terima kasih!"
            exit 0
            ;;
        *)
            echo ""
            echo "Pilihan tidak valid! Silakan pilih angka 0 sampai 4."
            sleep 2
            ;;
    esac
done
