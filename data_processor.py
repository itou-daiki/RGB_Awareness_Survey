
import pandas as pd
from config import ALL_QUESTIONS, SCORE_MAP

def preprocess_data(df):
    """
    A unified function to preprocess the raw survey data.
    - Cleans column names.
    - Extracts Grade (as int) and Class from the 4-digit ID.
    - Converts text-based survey answers to numerical scores.
    - Handles potential data errors.
    """
    df.columns = [c.strip() for c in df.columns]
    
    # Extract Grade and Class from the ID column
    id_col = "あなたのクラスと出席番号を4桁の数字で入力してください　例）1年6組34番 ⇒ 1634"
    if id_col in df.columns:
        # Ensure the column is treated as a string for manipulation
        id_str = df[id_col].astype(str).str.zfill(4)
        df['学年'] = pd.to_numeric(id_str.str[0], errors='coerce').astype('Int64') # Use Int64 to handle potential NaNs
        df['クラス'] = id_str.str[1] + "組"
    
    # Convert all question answers to numerical scores
    for q in ALL_QUESTIONS:
        if q in df.columns:
            df[q] = df[q].replace(SCORE_MAP)
            df[q] = pd.to_numeric(df[q], errors='coerce')
            
    return df
