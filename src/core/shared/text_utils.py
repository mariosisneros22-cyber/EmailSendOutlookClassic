import pandas as pd


def _cell_text(v) -> str:
    """
    Convierte valores de Excel a texto:
    - NaN/None => ""
    - strings => strip()
    - otros => str(...).strip()
    """
    if pd.isna(v):
        return ""
    
    return str(v).strip()