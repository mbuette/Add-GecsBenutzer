Laut Copilot, der aber nicht in der Lage war, die .md Datei zu erzeugen...

# 📘 Skript-Dokumentation: `Add-GecsBenutzer.ps1`

## 🔍 Zweck
Dieses PowerShell-Skript automatisiert die Erstellung von Benutzerkonten in einer Active Directory-Umgebung basierend auf einer CSV-Datei. Es legt bei Bedarf auch Organisationseinheiten (OUs) und Gruppen an und protokolliert alle Aktionen im Windows-Ereignisprotokoll.

---

## 📂 Eingabedatei

**Dateiname:** `GecsBenutzer.csv`  
**Format:** UTF-8 codiert, Spaltentrennung mit Komma oder Semikolon  

**Erforderliche Spalten:**
- `Vorname`
- `Nachname`
- `Standort`
- `Abteilung`
- `Passwort`

**Beispiel:**
```csv
Vorname,Nachname,Standort,Abteilung,Passwort
Max,Muster,Berlin,IT,Geheim123!
