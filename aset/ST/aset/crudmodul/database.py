import pandas as pd
import os

EXCEL_FILE = 'data.xlsx'

def load_data():
    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        dfs = {sheet: pd.read_excel(EXCEL_FILE, sheet_name=sheet) for sheet in xls.sheet_names}
    else:
        dfs = {
            'ST': pd.DataFrame(columns=['idST', 'Tgl', 'nost', 'lamahari']),
            'tim': pd.DataFrame(columns=['idtim', 'idST', 'tim', 'jl', 'kel', 'kec', 'kab', 's', 'e', 'pelaksanaan']),
            'anggota': pd.DataFrame(columns=['idang', 'idtim', 'namatim', 'nama', 'posisi'])
        }
    return dfs

def save_data(dfs):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')
