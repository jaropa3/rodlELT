import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
from config import SHEET_MAGAZYN, SHEET_SPRZEDAZ


def OPEN_file():
    # === Okno wyboru pliku ===
    root = Tk()
    root.withdraw()  # nie pokazuje pustego okna
    root.update()
    plik_wejsciowy = filedialog.askopenfilename()
    title="Wybierz plik Excel",
    filetypes=[("Pliki Excel", "*.xlsx *.xls")]

    root.destroy()

    if not plik_wejsciowy:
        raise SystemExit("Nie wybrano pliku.")

    return Path(plik_wejsciowy) # gdy jest return zawsze gdzieś przypisuj funcje, inaczej wartośc zostanie utracona

def read_input(file_path: Path):
    df_magazyn = pd.read_excel(file_path, sheet_name=SHEET_MAGAZYN)
    df_sprzedaz = pd.read_excel(file_path, sheet_name=SHEET_SPRZEDAZ)
    return df_magazyn, df_sprzedaz