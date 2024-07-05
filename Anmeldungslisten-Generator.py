# Input: Reporter-Abfrage "Angemeldete Teilnehmer (Daten für Anmeldezettel)"
#        in Zwischenablage, dann dieses Skript starten
# Output: CSV-Tabelle, die als Basis für eine Rundmail an Teilnehmer und deren Eltern
#        für eine bestimmte Veranstaltung genutzt werden kann inkl. Abfrage
#        wichtiger Kontaktdaten

import csv
import io
import win32clipboard
import datetime
from collections import defaultdict
from pprint import pprint


mit_eltern = False#True # Soll auch für die Eltern ein Datensatz erzeugt werden?
                  # (nur sinnvoll bei Versand per Mail)

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d_%H%M%S")

print("Bitte Reporter-Abfrage 'Angemeldete Teilnehmer (Daten für Anmeldezettel)'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()


if not data.startswith("lfd. Nr.\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

with io.StringIO(data) as infile:
    reader = csv.DictReader(infile, delimiter="\t")
    orig_feldnamen = reader.fieldnames
    daten = list(reader)

def komm_typ(item):
    if "@" in item: return ("E-Mail", item)
    if item.strip().startswith("01"): return ("Mobil", item)
    if item.strip().startswith("0"): return ("Festnetz", item)
    raise ValueError(f"Ungültiger Eintrag in {item}\n(Telefonnummern müssen mit 0 beginnen, Mails müssen ein @ enthalten)")

# Zunächst alle Daten systematisch zusammenfassen pro Kind,
# dabei Festnetz-, Mobilnummern und Mailadressen getrennt sammeln
teilnehmer = []
start = True
for eintrag in daten:
    if eintrag["lfd. Nr."]:
        if not start: 
            teilnehmer.append(datensatz)
        start = False
        datensatz = eintrag.copy()
        del datensatz["E-Mail_K"]
        datensatz["Komm_K"] = defaultdict(set) # alle Kontaktdaten des Kindes
        datensatz["Komm_E"] = defaultdict(set) # alle Kontaktdaten der Eltern
        datensatz["Mails"] = set()             # Mailadressen aller Personen
        datensatz["Festnetz"] = set()          # Festnetznummern aller Personen
        for gruppe in ("Komm_K", "Komm_E"):
            if eintrag[gruppe]:
                typ, nr = komm_typ(eintrag[gruppe])
                if typ == "E-Mail":
                    datensatz["Mails"].add(nr)
                if typ == "Festnetz":
                    datensatz["Festnetz"].add(nr)
                else:
                    datensatz[gruppe][typ].add(nr)

        if eintrag["E-Mail_K"]:
            datensatz["Komm_K"]["E-Mail"].add(eintrag["E-Mail_K"])
            datensatz["Mails"].add(eintrag["E-Mail_K"])
    else:
        for item in eintrag:
            if eintrag[item]:
                typ, nr = komm_typ(eintrag[item])
                if item == "E-Mail_K": item = "Komm_K"
                if typ == "E-Mail":
                    datensatz["Mails"].add(nr)
                if typ == "Festnetz":
                    datensatz["Festnetz"].add(nr)
                else:
                    datensatz[item][typ].add(nr)

teilnehmer.append(datensatz)

max_items={}

for gruppe in ("Komm_K", "Komm_E"):
    for typ in ("E-Mail", "Mobil"):
        max_items[gruppe, typ] = max(len(eintrag[gruppe][typ]) for eintrag in teilnehmer)

max_items["Festnetz"] = max(len(eintrag["Festnetz"]) for eintrag in teilnehmer)

feldnamen = []
for feld in orig_feldnamen:
    if feld.startswith("Komm_"):
        for typ in ("E-Mail", "Mobil"):
            for number in range(max_items[feld, typ]):
                feldnamen.append(f"{typ}_{feld[-1]}_{number+1}")
    elif feld != "E-Mail_K":
        feldnamen.append(feld)
for number in range(max_items["Festnetz"]):
    feldnamen.append(f"Festnetz_{number+1}")
feldnamen.append("Empfänger")

with open(f"Anmeldeliste_{heute}.csv", mode="w", newline="", encoding="cp1252") as outfile:
    output = csv.DictWriter(outfile, feldnamen, delimiter=";")
    output.writeheader()
    for person in teilnehmer:
        datensatz = {}
        for feld in orig_feldnamen:
            if feld.startswith("Komm_"):
                for typ in ("E-Mail", "Mobil"):
                    for number, eintrag in enumerate(person[feld][typ]):
                        datensatz[f"{typ}_{feld[-1]}_{number+1}"] = eintrag
            elif feld not in ("E-Mail_K", "Empfänger"):
                datensatz[feld] = person[feld]
        for number, eintrag in enumerate(person["Festnetz"]):
            datensatz[f"Festnetz_{number+1}"] = eintrag
        if mit_eltern:
            for empfänger in person["Mails"]:
                datensatz["Empfänger"] = empfänger
                output.writerow(datensatz)
        else:
            empfänger = next(iter(person["Mails"]))
            datensatz["Empfänger"] = empfänger
            output.writerow(datensatz)
            
        
input(f"Fertig! Die Datei Anmeldeliste_{heute}.csv wurde im aktuellen Ordner abgelegt.\nENTER drücken zum Beenden.")
