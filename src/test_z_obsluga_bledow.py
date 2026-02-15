import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
from transform import normalize_keys
from extract import OPEN_file


#def main():

wybrany_plik = OPEN_file()

plik_wyjsciowy = Path(r"Magazyn_wynik.xlsx")

wiersze_do_dodania = pd.read_excel

#print(wiersze_do_dodania)

try:
    
    # === Wczytanie arkuszy ===
    df_magazyn = pd.read_excel(wybrany_plik, sheet_name="Magazyn")
    df_sprzedaz = pd.read_excel(wybrany_plik, sheet_name="Sprzedane i zwrócone")

    # === Kolumna Liczba ===
    if "Liczba" not in df_magazyn.columns:
        df_magazyn.insert(10, "Liczba", "")
    else:
        df_magazyn["Liczba"] = ""

    # === Normalizacja kluczy ===
    df_magazyn = normalize_keys(df_magazyn)
    df_sprzedaz = normalize_keys(df_sprzedaz)
    # Filtrujemy wiersze, które chcemy dodać

    wiersze_do_dodania = df_sprzedaz[df_sprzedaz.iloc[:, 1] == -1].copy()
    wiersze_do_dodania = wiersze_do_dodania.reindex(columns=df_sprzedaz.columns, fill_value=None)#reindex(columns=...) Pandas Ustawia dokładnie taki układ kolumn jak w df_magazyn

    df_magazyn = pd.concat([df_magazyn, wiersze_do_dodania], ignore_index=True) # To jest łączenie dwóch DataFrame’ów... concat() zwraca nowy DataFrame. Dlatego: df_magazyn = ...bez tego wynik przepada... ignore_index=True — kluczowy parametr. bez duplikaty indeksów. Czyli reset indeksu. W 99% przypadków przy dokładaniu wierszy — chcesz to. 

    #przerobić na funkcje
    df_sprzedaz_1 = df_sprzedaz[df_sprzedaz.iloc[:, 1] == 1]
    
    # === Liczniki ===
    liczniki = df_sprzedaz_1["Key"].value_counts()

    # === Wpis tylko do ostatniego wystąpienia ===
    for key, count in liczniki.items():
        mask = df_magazyn["Key"] == key
        if mask.any():
            last_index = df_magazyn.loc[mask].index[-1]
            df_magazyn.loc[last_index, "Liczba"] = int(count)

    #/przerobić na funkcje
    
    # === Zapis pliku === Przerobić na funkcje
    df_magazyn.to_excel(plik_wyjsciowy, sheet_name="Magazyn_po_zmianach", index=False)
    messagebox.showinfo(
    title="Zakończono",
    message=f"✅ Plik został zapisany jako:\n\n{plik_wyjsciowy.absolute()}"
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

