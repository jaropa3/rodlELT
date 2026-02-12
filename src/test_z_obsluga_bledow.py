import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
#from transform import 
from extract import OPEN

# === Okno wyboru pliku ===
root = Tk()
root.withdraw()  # nie pokazuje pustego okna

# Ścieżki
plik_wejsciowy = filedialog.askopenfilename(
    title="Wybierz plik Excel",
    filetypes=[("Pliki Excel", "*.xlsx *.xls")]
)

if not plik_wejsciowy:
    raise SystemExit("❌ Nie wybrano pliku")

plik_wejsciowy = Path(plik_wejsciowy)
plik_wyjsciowy = Path(r"Magazyn_wynik.xlsx")

print("Wybrany plik:", plik_wejsciowy)

df_plik_wejsciowy = pd.read_excel(plik_wejsciowy)

wiersze_do_dodania = pd.read_excel

df_arkusz2_zwroty_zrodlowy = pd.read_excel(plik_wejsciowy, sheet_name="Sprzedane i zwrócone")

# Filtrujemy wiersze, które chcemy dodać
wiersze_do_dodania = df_arkusz2_zwroty_zrodlowy[df_arkusz2_zwroty_zrodlowy.iloc[:, 1] == -1]
wiersze_do_dodania.insert(6,"kolumna7", None)
wiersze_do_dodania.insert(7,"kolumna8", None)
wiersze_do_dodania.insert(8,"kolumna9", None)
wiersze_do_dodania.insert(9,"kolumna10", None)
wiersze_do_dodania.insert(10,"kolumna11", None)


#print(wiersze_do_dodania)

try:
    # === Sprawdzenie pliku ===
    if not plik_wejsciowy.exists():
        raise FileNotFoundError(f"Brak pliku wejściowego: {plik_wejsciowy}")

    # === Wczytanie arkuszy ===
    df_magazyn = pd.read_excel(plik_wejsciowy, sheet_name="Magazyn")
    df_sprzedaz = pd.read_excel(plik_wejsciowy, sheet_name="Sprzedane i zwrócone")

    # === Kolumna Liczba ===
    if "Liczba" not in df_magazyn.columns:
        df_magazyn.insert(10, "Liczba", "")
    else:
        df_magazyn["Liczba"] = ""

    # === Normalizacja kluczy ===
    df_magazyn["Key"] = df_magazyn["Key"].astype(str).str.strip()
    df_sprzedaz["Key"] = df_sprzedaz["Key"].astype(str).str.strip()

    df_sprzedaz_1 = df_sprzedaz[df_sprzedaz.iloc[:, 1] == 1]
    
    # === Liczniki ===
    liczniki = df_sprzedaz_1["Key"].value_counts()

    # === Wpis tylko do ostatniego wystąpienia ===
    for key, count in liczniki.items():
        mask = df_magazyn["Key"] == key
        if mask.any():
            last_index = df_magazyn.loc[mask].index[-1]
            df_magazyn.loc[last_index, "Liczba"] = int(count)

    for row in wiersze_do_dodania.itertuples(index=False):
        df_magazyn.loc[len(df_magazyn)] = list(row)  # konwertujemy tuple na listę

    # === Zapis pliku ===
    df_magazyn.to_excel(plik_wyjsciowy, sheet_name="Magazyn_po_zmianach", index=False)
    messagebox.showinfo(
    title="Zakończono",
    message=f"✅ Plik został zapisany jako:\n\n{plik_wyjsciowy.absolute}"
)



except FileNotFoundError as e:
    print(f"❌ {e}")
    sys.exit(1)

except PermissionError:
    messagebox.showinfo(
        title="Błąd",
    message=f"❌ Plik wyjściowy jest otwarty: {plik_wyjsciowy}"
)
    sys.exit(1)

except Exception as e:
    print("❌ Nieoczekiwany błąd:", type(e).__name__, e)
    sys.exit(1)

