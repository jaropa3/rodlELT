import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path

def OPEN_file():
    # === Okno wyboru pliku ===
    root = Tk()
    root.withdraw()  # nie pokazuje pustego okna
    plik_wejsciowy = filedialog.askopenfilename()
    title="Wybierz plik Excel",
    filetypes=[("Pliki Excel", "*.xlsx *.xls")]

    root.destroy()

    if not plik_wejsciowy:
        raise SystemExit("Nie wybrano pliku.")

    return Path(plik_wejsciowy) # gdy jest return zawsze gdzieś przypisuj funcje, inaczej wartośc zostanie utracona
