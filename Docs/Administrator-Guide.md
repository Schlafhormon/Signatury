# Signatury Administrator-Anleitung

## Vorlagen ablegen

- Für den lokalen Prototyp liegt die Vorlage unter `Templates\DK_Standard.docx`.
- Für den späteren Rollout werden die Vorlagen zentral unter `\\fileserver\Signatury\Templates` abgelegt.
- Die Dateinamen aus der Konfiguration müssen exakt mit den realen DOCX-Dateien übereinstimmen.

## Vorlagen aktualisieren

- Bestehende DOCX auf der Vorlagenquelle ersetzen.
- Beim nächsten Lauf erkennt das Skript die Änderung über den SHA256-Hash und spiegelt die neue Version in den lokalen Cache.
- Eine manuelle Bereinigung am Client ist im Normalfall nicht nötig.

## Neue Variablen ergänzen

1. Das gewünschte AD-Attribut festlegen.
2. In der DOCX einen neuen Platzhalter im Format `$MeinPlatzhalter$` verwenden.
3. In `Config\signature-config.json` unter `VariableMappings` ergänzen.
4. Skript erneut ausführen.

Beispiel:

```json
"$CurrentUserDepartment$": "department"
```

## Nur lokal testen

Lokal ohne Profiländerung:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -SkipDefaultSignatureUpdate
```

Lokal mit erzwungener Neuerzeugung:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -ForceRefresh
```

Lokal ohne AD-Verbindung:

- In `VariableOverrides` testweise Werte eintragen.
- Das Skript nutzt diese Werte, wenn AD leer bleibt oder nicht erreichbar ist.

## Empfohlene zentrale Struktur

```text
\\fileserver\Signatury\
  Templates\
    DK_Standard.docx
  Config\
    signature-config.json
  Scripts\
    Deploy-OutlookSignature.ps1
    Invoke-OutlookSignatureRollout.ps1
    OutlookSignature.Core.psm1
```

## GPO-Login-Beispiel

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "\\fileserver\Signatury\Scripts\Invoke-OutlookSignatureRollout.ps1" -ConfigPath "\\fileserver\Signatury\Config\signature-config.json"
```
