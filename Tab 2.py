import pandas as pd
import matplotlib.pyplot as plt
import re

# ----------------------------
# Einstellungen
# ----------------------------
XLSX_PATH = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'   # Pfad anpassen
TOTAL_VERSE = 1693                        # Anzahl Verse im Drama
WINDOW_SIZE = 50                       # Fensterbreite für Mittelwert (anpassen)

# ----------------------------
# Daten einlesen
# ----------------------------
df = pd.read_excel(XLSX_PATH)
or_col = df.iloc[:, 0].dropna().astype(str)

# ----------------------------
# Versnummern (vor dem Punkt) extrahieren
# ----------------------------
verses = []
for s in or_col:
    m = re.search(r"\s*(\d+)\.", s, flags=re.IGNORECASE)
    if m:
        verses.append(int(m.group(1)))

# Nur eindeutige Verse behalten
unique_verses = sorted(set(verses))

# ----------------------------
# Array: 1 wenn kommentiert, sonst 0
# ----------------------------
y_vals = [1 if v in unique_verses else 0 for v in range(1, TOTAL_VERSE + 1)]
x_vals = [v / TOTAL_VERSE for v in range(1, TOTAL_VERSE + 1)]

# ----------------------------
# Gleitender Mittelwert (zentriertes Fenster)
# ----------------------------
def moving_average(data, window):
    if window < 1:
        return data
    out = []
    half = window // 2
    for i in range(len(data)):
        start = max(0, i - half)
        end = min(len(data), i + half + 1)
        out.append(sum(data[start:end]) / (end - start))
    return out

y_smooth = moving_average(y_vals, WINDOW_SIZE)

# ----------------------------
# Plot
# ----------------------------
plt.figure(figsize=(12, 6))
plt.plot(x_vals, y_vals, color="steelblue", alpha=0.4, linewidth=1, label="Kommentierte Verse (1 = Kommentar)")
plt.plot(x_vals, y_smooth, color="darkred", linewidth=2, label=f"Gleitender Mittelwert (Fenster = {WINDOW_SIZE})")

plt.axhline(0, color="black", linewidth=1)
plt.title("Tab. 2: Verteilung der kommentierten Verse (1 = Kommentar vorhanden)")
plt.xlabel("Dramenverlauf in Prozent")
plt.ylabel("Kommentar vorhanden (1) / nicht (0)")
plt.grid(True, linestyle="--", alpha=0.6)
plt.legend()
plt.tight_layout()
plt.show()

# ----------------------------
# Statistik ausgeben
# ----------------------------
print(f"Kommentierte Verse: {len(unique_verses)} von {TOTAL_VERSE} ({len(unique_verses)/TOTAL_VERSE*100:.1f}%)")
