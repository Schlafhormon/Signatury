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

## Outlook-Defaultwerte feinsteuern

Seit der aktuellen Version können die Outlook-Schreibpfade getrennt gesteuert werden:

- `WriteDefaultsToMailSettings`: schreibt `NewSignature` und `ReplySignature` nach `HKCU\Software\Microsoft\Office\<Version>\Common\MailSettings`
- `WriteDefaultsToProfileAccounts`: schreibt `New Signature` und `Reply-Forward Signature` zusätzlich in das erkannte Outlook-Profil

Empfehlung für Umgebungen, in denen der Outlook-Signaturdialog nach dem Rollout nicht mehr bearbeitbar ist:

```json
"WriteDefaultsToMailSettings": true,
"WriteDefaultsToProfileAccounts": false,
"DisableRoamingSignatures": false
```

Damit werden die Signaturdateien weiterhin erzeugt und die Standardsignatur in `Common\MailSettings` gesetzt, aber die aggressiveren Profil-Account-Änderungen bleiben aus.

## Verwalteten Zustand vor einem Lauf zurücksetzen

Mit `ResetManagedStateBeforeDeploy` kann Signatury vor dem eigentlichen Lauf einen verwalteten "Frischstart" erzwingen.

```json
"ResetManagedStateBeforeDeploy": true
```

Dabei entfernt Signatury vor dem neuen Lauf:

- bereits erzeugte Signaturdateien der konfigurierten Signaturnamen
- zugehörige Asset-Ordner
- Statusdateien für Fingerprints
- lokale Cache-Kopien der betroffenen Vorlagen
- von Signatury gesetzte Outlook-Defaultwerte in `Common\MailSettings`
- von Signatury gesetzte `New Signature`- und `Reply-Forward Signature`-Werte im Outlook-Profil
- `DisableRoamingSignatures`-Schalter im Outlook-Setup

Das ist besonders hilfreich als Recovery-Option nach einem fehlerhaften Rollout. Solange der Schalter aktiv ist, wird allerdings bei jedem Lauf bewusst ein vollständiger Neuaufbau erzwungen.

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
