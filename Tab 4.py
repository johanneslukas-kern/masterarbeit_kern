import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
import re

# ----------------------------
# Einstellungen
# ----------------------------
XLSX_PATH = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'  # Pfad ggf. anpassen

# ----------------------------
# Excel einlesen
# ----------------------------
df = pd.read_excel(XLSX_PATH)

# Spalte B enthält Personennamen oder Rollen
if df.shape[1] < 2:
    raise ValueError("In der Excel-Datei gibt es keine zweite Spalte mit Personenangaben.")

person_col = df.iloc[:, 1].dropna().astype(str).str.strip()

# ----------------------------
# Mehrere Namen pro Zelle aufsplitten
# ----------------------------
# Trennzeichen können z.B. Komma, Semikolon oder Leerzeichen sein
all_names = []
for cell in person_col:
    # Zerlegen in einzelne Wörter/Namen
    names = re.split(r"[,\s;/&]+", cell.strip())
    # Nur nicht-leere Namen übernehmen
    names = [n for n in names if n]
    all_names.extend(names)

# ----------------------------
# Häufigkeiten zählen
# ----------------------------
counter = Counter(all_names)
sorted_items = sorted(counter.items(), key=lambda x: x[1], reverse=True)

labels = [item[0] for item in sorted_items]
counts = [item[1] for item in sorted_items]

# ----------------------------
# Balkendiagramm erstellen
# ----------------------------
plt.figure(figsize=(12, 6))
plt.bar(labels, counts, color="teal")
plt.title("Tab. 4: Anzahl der Kommentare pro Figur")
plt.xlabel("Figur")
plt.ylabel("Anzahl der Kommentare")
plt.xticks(rotation=45, ha="right")
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()

# ----------------------------
# Zusatz: Top 10 Personen ausgeben
# ----------------------------
print("Top 10 Personen mit den meisten Kommentaren:")
for name, count in sorted_items[:10]:
    print(f"{name}: {count}")