import pandas as pd
import sys
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
from transform import normalize_keys, count_apply, prepare_magazyn
from extract import OPEN_file, read_input


def main():

    wybrany_plik = OPEN_file()
    plik_wyjsciowy = Path(r"Magazyn_wynik.xlsx")

    wiersze_do_dodania = pd.read_excel

    #print(wiersze_do_dodania)

    try:
        
        # === Wczytanie arkuszy === EXTRACT 
        df_magazyn, df_sprzedaz = read_input(wybrany_plik)
               
        # === TRANSFORM ===

        df_magazyn = prepare_magazyn(df_magazyn)

        # === Normalizacja kluczy ===
        df_magazyn = normalize_keys(df_magazyn)
        df_sprzedaz = normalize_keys(df_sprzedaz)

        # Filtrujemy wiersze, które chcemy dodać

        wiersze_do_dodania = df_sprzedaz[df_sprzedaz.iloc[:, 1] == -1].copy()
        wiersze_do_dodania = wiersze_do_dodania.reindex(columns=df_sprzedaz.columns, fill_value=None)#reindex(columns=...) Pandas Ustawia dokładnie taki układ kolumn jak w df_magazyn

        df_magazyn = pd.concat([df_magazyn, wiersze_do_dodania], ignore_index=True) # To jest łączenie dwóch DataFrame’ów... concat() zwraca nowy DataFrame. Dlatego: df_magazyn = ...bez tego wynik przepada... ignore_index=True — kluczowy parametr. bez duplikaty indeksów. Czyli reset indeksu. W 99% przypadków przy dokładaniu wierszy — chcesz to. 
        df_magazyn = count_apply(df_magazyn, df_sprzedaz)
        
        # === Zapis pliku === Przerobić na funkcje === LOAD
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
        messagebox.showerror("Błąd", str(e))

if __name__ == "__main__":
    main()