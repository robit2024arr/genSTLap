#!/bin/bash

while true; do
    clear
    echo "=============================="
    echo "       MENU APLIKASI CRUD     "
    echo "=============================="
    echo "1. Jalankan Aplikasi CRUD Python (data.xlsx)"
    echo "2. Jalankan Generator (generator.sh)"
    echo "3. Keluar"
    echo "=============================="
    read -p "Pilih menu [1-3]: " pilihan

    case $pilihan in
        1)
            echo ""
            python3 aset/crud.py
            echo ""
            read -p "Tekan [Enter] untuk kembali ke menu..."
            ;;
        2)
            echo ""
            bash aset/generator.sh
            echo ""
            read -p "Tekan [Enter] untuk kembali ke menu..."
            ;;
        3)
            echo "Keluar dari program."
            exit 0
            ;;
        *)
            echo ""
            echo "Pilihan tidak valid. Silakan pilih 1-3."
            sleep 1
            ;;
    esac
done

```
