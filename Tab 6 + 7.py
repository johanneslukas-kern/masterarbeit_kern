import re
import pandas as pd
import matplotlib.pyplot as plt
from nltk.sentiment import SentimentIntensityAnalyzer
import numpy as np
from collections import defaultdict
import math

# --- Einstellungen ---
XLSX_PATH = '/Users/johannes/Desktop/Masterarbeit Sammlung/Export_Textsammlung.xlsx'
TOTAL_VERSE = 1693
WINDOW_SIZE = 100  # Fensterbreite gleitender Mittelwert

# --- Daten einlesen ---
df = pd.read_excel(XLSX_PATH)

# Übersetzung fix in Spalte E (Index 4)
text_col = df.columns[4]
print(f"Analyse-Spalte: {text_col}")

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

# --- Fehlende Werte "halten" statt auf 0 zurückfallen ---
y_series = pd.Series(y_per_verse)
y_filled = y_series.fillna(method='ffill').fillna(method='bfill')  # ffill: letzter bekannter Wert

# --- Gleitender Mittelwert ---
y_smooth = y_filled.rolling(window=WINDOW_SIZE, center=True, min_periods=1).mean()
x_line = np.arange(1, TOTAL_VERSE + 1) / TOTAL_VERSE

# --- Statistische Kennzahlen ---
sent = df["Sentiment"].dropna()
mean_sent = sent.mean()
median_sent = sent.median()
std_sent = sent.std()
n = len(sent)

print("\n--- Sentiment-Statistik ---")
print(f"Anzahl analysierter Kommentare: {n}")
print(f"Mittelwert (compound): {mean_sent:.4f}")
print(f"Median: {median_sent:.4f}")
print(f"Standardabweichung: {std_sent:.4f}")

# Einfacher t-Wert (ungefähr)
if n > 1:
    t_stat = mean_sent / (std_sent / math.sqrt(n))
    print(f"T-Wert (ungefähr): {t_stat:.3f}")
    if abs(t_stat) > 1.96:
        if mean_sent > 0:
            print("⇒ Durchschnittliches Sentiment ist vermutlich POSITIV (p < 0.05).")
        else:
            print("⇒ Durchschnittliches Sentiment ist vermutlich NEGATIV (p < 0.05).")
    else:
        print("⇒ Keine signifikante Abweichung von neutral (p ≥ 0.05, Näherung).")

# --- Scatter + Gleitender Mittelwert + Medianlinie ---
plt.figure(figsize=(12, 6))
plt.scatter(df["Verse"]/TOTAL_VERSE, df["Sentiment"], s=12, alpha=0.4, label="Kommentare")
plt.plot(x_line, y_smooth, color="darkred", linewidth=2,
         label=f"Gleitender Mittelwert (w={WINDOW_SIZE})")
plt.axhline(0, color="black", linewidth=1, linestyle="--", label="Neutral (0)")
plt.axhline(median_sent, color="orange", linewidth=1.8, linestyle="--",
            label=f"Median = {median_sent:.3f}")


plt.title("Tab. 6: Sentiment der englischen Kommentare")
plt.xlabel("Dramenverlauf in Prozent")
plt.ylabel("Sentiment (−1 … +1)")
plt.grid(True, linestyle="--", alpha=0.5)
plt.legend()
plt.tight_layout()
plt.show()

from nltk.sentiment.vader import SentimentIntensityAnalyzer
from collections import Counter
import re


# --- VADER Lexikon ---
sia = SentimentIntensityAnalyzer()
vader_lex = sia.lexicon  # Wörterbuch: Wort -> Score

# --- Alle Wörter aus deinen Kommentaren sammeln ---
all_comments = df['EN'].tolist()
found_words = []

for comment in all_comments:
    words = re.findall(r'\b\w+\b', comment.lower())
    for w in words:
        if w in vader_lex:  # nur Wörter, die bewertet werden
            found_words.append((w, vader_lex[w]))

# --- In Counter / DataFrame zusammenfassen ---
found_counts = Counter([w for w, score in found_words])
found_df = pd.DataFrame(found_words, columns=['Wort', 'Score'])
found_df['Häufigkeit'] = found_df['Wort'].map(found_counts)

# --- Nach Häufigkeit sortieren ---
found_df = found_df.drop_duplicates(subset=['Wort']).sort_values(by='Häufigkeit', ascending=False)
pos = (sent > 0).sum()
neg = (sent < 0).sum()
n_nonzero = pos + neg
from math import comb

# --- Kategorien erstellen ---
sent_cat = df['Sentiment'].apply(
    lambda s: 'Positiv' if s > 0.05 else ('Negativ' if s < -0.05 else 'Neutral')
)

# --- Häufigkeit zählen ---
sent_counts = sent_cat.value_counts()

print("\n--- Anzahl der Kommentare nach Sentiment ---")
print(f"Positiv: {sent_counts.get('Positiv', 0)}")
print(f"Neutral: {sent_counts.get('Neutral', 0)}")
print(f"Negativ: {sent_counts.get('Negativ', 0)}")

print("Bewertete Wörter in deinen Texten:")
print(found_df.head(30).to_string(index=False))  # Top 30 Wörter
# --- Tortendiagramm: Positiv / Neutral / Negativ ---
sent_cat = df.apply(lambda row: 'Positiv' if row['Sentiment'] > 0.05 
                    else ('Negativ' if row['Sentiment'] < -0.05 else 'Neutral'), axis=1)
counts = sent_cat.value_counts()
plt.figure(figsize=(6,6))
plt.pie(counts, labels=counts.index, autopct='%1.1f%%', colors=['red','green','grey'])
plt.title("Tab. 7: Verteilung der Sentimente")
plt.show()