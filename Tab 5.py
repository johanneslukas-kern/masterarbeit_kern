import xml.etree.ElementTree as ET
import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter

# ----------------------------
# Einstellungen
# ----------------------------
XML_PATH = "/Users/johannes/Desktop/Masterarbeit Sammlung/Orestes.xml"  # Pfad anpassen
TOP_N = 10  # Nur die 10 wichtigsten Figuren

# ----------------------------
# XML einlesen
# ----------------------------
tree = ET.parse(XML_PATH)
root = tree.getroot()

# TEI-Namespace erkennen (falls vorhanden)
ns = {"tei": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}

# ----------------------------
# Sprecher und Verse zählen
# ----------------------------
speaker_counts = Counter()

# Alle <sp>-Elemente durchlaufen
for sp in root.findall(".//tei:sp", ns) if ns else root.findall(".//sp"):
    speaker_elem = sp.find("tei:speaker", ns) if ns else sp.find("speaker")
    if speaker_elem is not None and speaker_elem.text:
        speaker = speaker_elem.text.strip()
        verses = sp.findall("tei:l", ns) if ns else sp.findall("l")
        speaker_counts[speaker] += len(verses)

# ----------------------------
# DataFrame erstellen und Top 10 auswählen
# ----------------------------
df = pd.DataFrame(list(speaker_counts.items()), columns=["Speaker", "Verse_Count"])
df = df.sort_values("Verse_Count", ascending=False).head(TOP_N)

# ----------------------------
# Statistik ausgeben
# ----------------------------
if df.empty:
    print("⚠️ Keine Sprecher gefunden. Prüfe, ob die XML-Struktur TEI-kompatibel ist.")
else:
    print(f"=== Top {TOP_N} Figuren nach Versanzahl ===")
    print(df.to_string(index=False))
    total = df["Verse_Count"].sum()
    print(f"\nGesamtzahl Verse dieser {TOP_N} Figuren: {total}")

    # ----------------------------
    # Plot
    # ----------------------------
    plt.figure(figsize=(10, 6))
    plt.bar(df["Speaker"], df["Verse_Count"], color="darkred")
    plt.title(f"Tab. 5: Figuren nach Versanzahl")
    plt.xlabel("Figur")
    plt.ylabel("Anzahl der Verse")
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()
    plt.show()