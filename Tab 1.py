import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
import re

# ----------------------------
# Einstellungen
# ----------------------------
XLSX_PATH   = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'  # Pfad ggf. anpassen
TOTAL_VERSE = 1693                         # Gesamtzahl Verse
WINDOW_SIZE = 30                           # Fensterbreite gleitender Mittelwert

# ----------------------------
# Daten einlesen: erste Spalte "Or. xxx.xx"
# ----------------------------
df = pd.read_excel(XLSX_PATH)
or_col = df.iloc[:, 0].dropna().astype(str)

# Versnummer VOR dem Punkt extrahieren (xxx aus "Or. xxx.xx")
verses = []
for s in or_col:
    m = re.search(r"\s*(\d+)\.", s, flags=re.IGNORECASE)
    if m:
        verses.append(int(m.group(1)))

# Häufigkeit pro Vers zählen und auf 1..TOTAL_VERSE auffüllen (Lücken = 0)
counts = Counter(verses)
y_counts = [counts.get(v, 0) for v in range(1, TOTAL_VERSE + 1)]
x_vals   = [v / TOTAL_VERSE for v in range(1, TOTAL_VERSE + 1)]

# Gleitender Mittelwert (zentriertes Fenster, aber NICHT um 0 verschoben)
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
# Plot: 0 als Baseline; nur bei Kommentaren geht die Kurve hoch
# ----------------------------
plt.figure(figsize=(12, 6))
plt.plot(x_vals, y_counts, label="Kommentare pro Vers", color="steelblue", linewidth=1)
plt.plot(x_vals, y_smooth, label=f"Gleitender Mittelwert (w={WINDOW_SIZE})", color="darkred", linewidth=2)
plt.axhline(0, color="black", linewidth=1)  # Baseline bei 0

plt.title("Tab. 1: Alle Kommentare pro Vers über den Dramenverlauf")
plt.xlabel("Dramenverlauf in Prozent")
plt.ylabel("Anzahl Kommentare")
plt.grid(True, linestyle="--", alpha=0.6)
plt.legend()
plt.tight_layout()
plt.show()
