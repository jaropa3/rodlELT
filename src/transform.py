import pandas as pd

def normalize_keys(df: pd.DataFrame) -> pd.DataFrame:
    df["Key"] = df["Key"].astype(str).str.strip()
    return df

def count_apply(df_sprzedaz, df_magazyn):
    
    df_sprzedane = df_sprzedaz[df_sprzedaz.iloc[:, 1] == 1]
    df_liczniki = df_sprzedane["Key"].value_counts()


    for key, count in df_liczniki.items():
        mask = df_magazyn["Key"] == key
        if mask.any():
            last_index = df_magazyn.loc[mask].index[-1]
            df_magazyn.loc[last_index, "Liczba"] = int(count)
    
    return df_magazyn

def prepare_magazyn(df_magazyn):
    if "Liczba" not in df_magazyn.columns:
        df_magazyn.insert(10, "Liczba", "")
    else:
        df_magazyn["Liczba"] = ""
    return df_magazyn