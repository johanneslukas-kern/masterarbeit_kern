import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
import re

# ----------------------------
# Einstellungen
# ----------------------------
XLSX_PATH   = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'  # Pfad ggf. anpassen
TOTAL_VERSE = 1693                         # Gesamtzahl Verse
WINDOW_SIZE = 50                          # Fensterbreite gleitender Mittelwert

# ----------------------------
# Daten einlesen
# ----------------------------
df = pd.read_excel(XLSX_PATH)

# ----------------------------
# Nur Zeilen behalten, in denen Spalte C "(vet exeg)" oder "(vet paraphr)" enthält
# ----------------------------
col_c = df.columns[2]  # Spalte C hat Index 2 (A=0, B=1, C=2)
df = df[df[col_c].astype(str).str.contains(r"\(vet\s+(exeg|paraphr)\)", flags=re.IGNORECASE, na=False)]

print(f"Gefilterte Zeilen: {len(df)}")

# ----------------------------
# Erste Spalte enthält Versangaben (z. B. '4.01', '5.02' usw.)
# ----------------------------
or_col = df.iloc[:, 0].dropna().astype(str)

# Versnummer VOR dem Punkt extrahieren (xxx aus '4.01' usw.)
verses = []
for s in or_col:
    m = re.search(r"\s*(\d+)\.", s, flags=re.IGNORECASE)
    if m:
        verses.append(int(m.group(1)))

# ----------------------------
# Häufigkeit pro Vers zählen und auf 1..TOTAL_VERSE auffüllen (Lücken = 0)
# ----------------------------
counts = Counter(verses)
y_counts = [counts.get(v, 0) for v in range(1, TOTAL_VERSE + 1)]
x_vals   = [v / TOTAL_VERSE for v in range(1, TOTAL_VERSE + 1)]

# ----------------------------
# Gleitender Mittelwert (zentriertes Fenster)
# ----------------------------
def moving_average(seq, w):
    if w < 2:
        return seq[:]
    half = w // 2
    out = []
    for i in range(len(seq)):
        a = max(0, i - half)
        b = min(len(seq), i + half + 1)
        out.append(sum(seq[a:b]) / (b - a))
    return out

y_smooth = moving_average(y_counts, WINDOW_SIZE)

# ----------------------------
# Plot
# ----------------------------
plt.figure(figsize=(12, 6))
plt.plot(x_vals, y_counts, label="Kommentare pro Vers", color="steelblue", linewidth=1)
plt.plot(x_vals, y_smooth, label=f"Gleitender Mittelwert (w={WINDOW_SIZE})", color="darkred", linewidth=2)
plt.axhline(0, color="black", linewidth=1)

plt.title("Tab. 3: Kommentare pro Vers über den Dramenverlauf (nur vet exeg / paraphr)")
plt.xlabel("Dramenverlauf in Prozent")
plt.ylabel("Anzahl Kommentare")
plt.grid(True, linestyle="--", alpha=0.6)
plt.legend()
plt.tight_layout()
plt.show()

# ----------------------------
# Statistik
# ----------------------------
print(f"Kommentierte Verse: {len(set(verses))} von {TOTAL_VERSE} "
      f"({len(set(verses))/TOTAL_VERSE*100:.1f}%)")