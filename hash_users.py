import pandas as pd
import bcrypt
from pathlib import Path

XLSX = Path("toegestaan_gebruik.xlsx")

df = pd.read_excel(XLSX, dtype=str).fillna("")
df.columns = [c.strip().lower() for c in df.columns]

required = {"naam", "rol", "paswoord"}
missing = required - set(df.columns)
if missing:
    raise ValueError(f"Ontbrekende kolommen: {missing}")

def hash_pw(pw: str) -> str:
    pw = pw.strip()
    if not pw:
        return ""
    return bcrypt.hashpw(pw.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

df["paswoord_hash"] = df["paswoord"].apply(hash_pw)

# Optioneel: wis plain text paswoorden (sterk aangeraden)
df["paswoord"] = ""

df.to_excel(XLSX, index=False)
print("Klaar. 'paswoord_hash' is toegevoegd en 'paswoord' is leeggemaakt.")
