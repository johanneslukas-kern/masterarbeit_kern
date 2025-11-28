import re
import pandas as pd
import matplotlib.pyplot as plt
from nltk.sentiment import SentimentIntensityAnalyzer
import numpy as np
from collections import defaultdict, Counter
import math

# --- Einstellungen ---
XLSX_PATH = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'
TOTAL_VERSE = 1693
WINDOW_SIZE = 30
DECAY_RATE = 0.4  # Anteil pro leeren Vers, um zum Nullpunkt zurückzugehen

# --- Daten einlesen ---
df = pd.read_excel(XLSX_PATH)

# --- Spalte für Kommentare finden ---
col_candidates = [c for c in df.columns if "trans" in str(c).lower()]
if col_candidates:
    text_col = col_candidates[0]
else:
    text_col = df.columns[-1]
print(f"Analyse-Spalte: {text_col}")

# --- Nur Zeilen von Menelaos behalten ---
speaker_col = df.columns[1]  # zweite Spalte
df = df[df[speaker_col].astype(str).str.contains(r"\bTyndareos\b", flags=re.IGNORECASE, na=False)]
print(f"Anzahl Zeilen mit Menelaos: {len(df)}")

# --- Versnummer extrahieren ---
def extract_verse(s):
    m = re.search(r"\s*(\d+)\.", str(s), flags=re.IGNORECASE)
    return int(m.group(1)) if m else None

df["Verse"] = df.iloc[:, 0].map(extract_verse)
df = df.dropna(subset=["Verse"])
df["EN"] = df[text_col].astype(str)
df = df[df["EN"].str.strip().ne("")]

# --- Sentiment berechnen ---
sia = SentimentIntensityAnalyzer()
df["Sentiment"] = df["EN"].apply(lambda t: sia.polarity_scores(t)["compound"])

# --- Pro Vers aggregieren ---
per_verse = defaultdict(list)
for v, s in zip(df["Verse"].astype(int), df["Sentiment"].astype(float)):
    if 1 <= v <= TOTAL_VERSE:
        per_verse[v].append(s)

y_per_verse = [np.mean(per_verse[v]) if v in per_verse else np.nan
               for v in range(1, TOTAL_VERSE + 1)]

# --- Fehlende Werte langsam zu 0 zurückführen ---
y_series = pd.Series(y_per_verse)
y_filled = []
current = 0.0
for val in y_series:
    if not np.isnan(val):
        current = val
    else:
        current = current * (1 - DECAY_RATE)
    y_filled.append(current)

# --- Gleitender Mittelwert ---
y_smooth = pd.Series(y_filled).rolling(window=WINDOW_SIZE, center=True, min_periods=1).mean()
x_line = np.arange(1, TOTAL_VERSE + 1) / TOTAL_VERSE

# --- Statistische Kennzahlen ---
sent = df["Sentiment"].dropna()
mean_sent = sent.mean()
median_sent = sent.median()
std_sent = sent.std()
n = len(sent)
t_stat = mean_sent / (std_sent / math.sqrt(n)) if n > 1 else float('nan')

# --- Sentiment-Kategorien ---
sent_cat = df.apply(lambda row: 'Positiv' if row['Sentiment'] > 0.05 
                    else ('Negativ' if row['Sentiment'] < -0.05 else 'Neutral'), axis=1)
counts = sent_cat.value_counts()

# --- Ausgabe ---
print("\n--- Sentiment-Statistik ---")
print(f"Anzahl analysierter Kommentare: {n}")
print(f"Mittelwert (compound): {mean_sent:.4f}")
print(f"Median: {median_sent:.4f}")
print(f"Standardabweichung: {std_sent:.4f}")
print(f"T-Wert (ungefähr): {t_stat:.3f}")
if abs(t_stat) > 1.96:
    if mean_sent > 0:
        print("⇒ Durchschnittliches Sentiment ist vermutlich POSITIV (p < 0.05).")
    else:
        print("⇒ Durchschnittliches Sentiment ist vermutlich NEGATIV (p < 0.05).")
else:
    print("⇒ Keine signifikante Abweichung von neutral (p ≥ 0.05, Näherung).")

print("\nAnzahl pro Sentiment-Kategorie:")
print(counts)

# --- Plot ---
plt.figure(figsize=(12, 6))
plt.scatter(df["Verse"]/TOTAL_VERSE, df["Sentiment"], s=12, alpha=0.4, label="Kommentare (Menelaos)")
plt.plot(x_line, y_smooth, color="darkred", linewidth=2, label=f"Gleitender Mittelwert (w={WINDOW_SIZE})")
plt.axhline(0, color="black", linewidth=1, linestyle="--", label="Neutral (0)")
plt.axhline(median_sent, color="orange", linewidth=1.8, linestyle="--", label=f"Median = {median_sent:.3f}")

plt.title("Tab. 8: Sentiment der englischen Kommentare zu Menelaos")
plt.xlabel("Dramenverlauf in Prozent")
plt.ylabel("Sentiment (−1 … +1)")
plt.grid(True, linestyle="--", alpha=0.5)
plt.legend()
plt.tight_layout()
plt.show()
