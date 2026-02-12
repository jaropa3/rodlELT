import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path

def OPEN():
    plik_wejsciowy = filedialog.askopenfilename(
    title="Wybierz plik Excel",
    filetypes=[("Pliki Excel", "*.xlsx *.xls")]
)