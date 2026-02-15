import pandas as pd

def normalize_keys(df: pd.DataFrame) -> pd.DataFrame:
    df["Key"] = df["Key"].astype(str).str.strip()
    return df
