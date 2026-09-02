import pandas as pd

def clean_float(val):
    """Converts raw cell value or pandas item into a float safely."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(',', '.')
    try:
        return float(val_str)
    except Exception:
        return 0.0

def normalize_columns(df):
    """
    Ensures dataframe columns are consistently labeled 'Colonna1', 'Colonna2', ...,
    even if the source Excel file was read with or without headers or has varying column counts.
    """
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=[f"Colonna{i+1}" for i in range(10)])
    if 'Colonna2' not in df.columns:
        row_data = pd.DataFrame([df.columns.values], columns=[f"Colonna{i+1}" for i in range(len(df.columns))])
        df.columns = [f"Colonna{i+1}" for i in range(len(df.columns))]
        df = pd.concat([row_data, df], ignore_index=True)
    return df
