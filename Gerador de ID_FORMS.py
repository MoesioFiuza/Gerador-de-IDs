import pandas as pd
import string
import random
import openpyxl
import tkinter as tk
from tkinter import filedialog

def generate_ids(file_path):
    df = pd.read_excel(file_path, sheet_name='aba novos ids', engine='openpyxl')

    new_ids = []
    existing_ids = df['ID_FORMULARIO'].dropna().unique().tolist()
    while len(new_ids) < df['ID_FORMULARIO'].isna().sum():
        id_chars = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
        if id_chars not in existing_ids:
            new_ids.append(id_chars)
            existing_ids.append(id_chars)

    df.loc[df['ID_FORMULARIO'].isna(), 'ID_FORMULARIO'] = new_ids

    new_file_name = file_path.replace('.xlsx', '_novo.xlsx')
    df.to_excel(new_file_name, index=False, engine='openpyxl')
    print(f"Arquivo Salvo em: {new_file_name}")

def select_file():
    root = tk.Tk()
    root.withdraw() 
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        generate_ids(file_path)

if __name__ == "__main__":
    select_file()