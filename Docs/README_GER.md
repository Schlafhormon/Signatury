# Signatury

Signatury ist eine PowerShell-basierte Outlook-Signaturlösung für Windows- und Active-Directory-Umgebungen. Die Lösung verwendet vorhandene Word-`.docx`-Vorlagen als Master, liest Benutzerdaten direkt aus Active Directory aus, erzeugt Outlook-kompatible Signaturdateien und setzt die Standardsignatur für Classic Outlook unter Windows.

Der aktuelle Stand ist lokal testbar, die Architektur ist aber bereits für einen späteren Rollout per GPO-Login-Skript, geplanter Aufgabe oder Intune im Benutzerkontext vorbereitet.

## Architektur

Die Lösung verwendet die vorhandene DOCX-Datei als Master und verarbeitet sie lokal im Benutzerkontext mit Microsoft Word per COM-Automation. Dadurch bleiben Formatierung, Tabellen, Bilder und Logos erhalten, während Platzhalter robust per Word-`Find.Execute` ersetzt werden, auch wenn Word den Text intern in mehrere Runs zerlegt.

Die Architektur trennt sauber zwischen Konfiguration, AD-Datenbeschaffung, lokalem Cache, Dokumentverarbeitung, Signaturausgabe und Outlook-Registry-Setzung. Für den aktuellen Prototyp laufen alle Pfade lokal im Repository beziehungsweise im Benutzerprofil. Für den späteren Unternehmensrollout müssen im Regelfall nur die Pfade in der JSON-Konfiguration auf UNC- und `%LOCALAPPDATA%`-Pfade umgestellt und das Startskript per GPO, geplanter Aufgabe oder Intune verteilt werden.

## Ordnerstruktur

```text
Templates\
  DK_Standard.docx
Config\
  signature-config.json
  signature-config.enterprise.example.json
Scripts\
  OutlookSignature.Core.psm1
  Deploy-OutlookSignature.ps1
  Invoke-OutlookSignatureRollout.ps1
Tests\
  Invoke-OutlookSignatureSelfTest.ps1
Docs\
  Administrator-Guide.md
  README_GER.md
Logs\
Temp\
  Cache\
  State\
  Work\
```

## Funktionsumfang

- Liest DOCX-Vorlagen aus einem konfigurierbaren Vorlagenordner
- Spiegelt Vorlagen in einen lokalen Cache und arbeitet bevorzugt aus diesem Cache
- Liest AD-Attribute des aktuell angemeldeten Domänenbenutzers
- Ersetzt Platzhalter in der DOCX robust über Word-COM
- Exportiert Outlook-kompatible `.htm`, `.txt` und optional `.rtf`
- Übernimmt eingebettete Bilder und Logos über den Word-HTML-Export
- Setzt Standard-Signaturen für neue Nachrichten und Antworten/Weiterleitungen per Registry
- Schreibt verständliche Benutzer-Logdateien
- Arbeitet idempotent und erzeugt nur neu, wenn sich Vorlage, Variablen oder relevante Konfiguration geändert haben

## Unterstützte Versionen

- Windows 10 und Windows 11
- Windows PowerShell 5.1
- Classic Outlook für Windows
- Office 2016, 2019, 2021 und Microsoft 365 Apps mit Outlook unter dem Registry-Zweig `Office\<Version>\Outlook`
- Word muss lokal installiert sein, da die DOCX-Verarbeitung per Word-COM erfolgt

## Voraussetzungen

- Benutzerkontext ohne lokale Administratorrechte
- Domänenanmeldung für produktiven AD-Betrieb
- Zugriff auf Active Directory aus dem Benutzerkontext
- Zugriff auf die Vorlagenquelle
- Microsoft Word installiert
- Classic Outlook installiert

## Konfiguration

Die Datei [Config/signature-config.json](../Config/signature-config.json) enthält alle relevanten Einstellungen.

Wichtige Konfigurationsschlüssel:

- `TemplateRootPath`: Vorlagenquelle, lokal oder UNC
- `CachePath`: lokaler Cache
- `LogPath`: Logdateien
- `TempPath`: temporäre Arbeitskopien
- `StatePath`: Fingerprint-Statusdateien für Idempotenz
- `SignatureOutputPath`: Outlook-Signaturordner
- `TemplateFileName`: Standardvorlage
- `SignatureName`: Standardname der Outlook-Signatur
- `WriteDefaultsToMailSettings`: schreibt Outlook-Standardwerte nach `Common\MailSettings`
- `WriteDefaultsToProfileAccounts`: schreibt Outlook-Standardwerte zusätzlich direkt in die Profil-Kontenschlüssel
- `ResetManagedStateBeforeDeploy`: setzt von Signatury verwaltete Dateien und Outlook-Defaultwerte vor dem Lauf zurück
- `VariableMappings`: Mapping Platzhalter zu AD-Attributen
- `VariableOverrides`: optionale lokale Fallbackwerte für Offline-Tests
- `OptionalLineCleanupPatterns`: Regex-Muster für das Entfernen leerer oder halbleerer Zeilen
- `Templates`: optionaler Einstieg für mehrere Vorlagen

Relative Pfade wie `.\Templates` werden bewusst relativ zum Projektstamm und nicht relativ zur JSON-Datei aufgelöst.

Wenn Outlook nach einem Rollout die Signaturverwaltung nicht mehr editierbar anzeigt, empfiehlt sich häufig folgende Konfiguration:

```json
"WriteDefaultsToMailSettings": true,
"WriteDefaultsToProfileAccounts": false,
"DisableRoamingSignatures": false
```

Damit setzt Signatury die Standardsignatur nur noch über `Common\MailSettings` und greift nicht mehr in die kontospezifischen Profilwerte ein.

Wenn ein vorheriger Rollout Outlook in einen inkonsistenten Zustand gebracht hat, kann zusätzlich folgender Schalter verwendet werden:

```json
"ResetManagedStateBeforeDeploy": true
```

Damit entfernt Signatury vor dem neuen Lauf seine verwalteten Signaturdateien, Statusdateien, Cache-Kopien und Outlook-Defaultwerte und verhält sich dadurch wie ein bewusst erzwungener Frischstart.

## Platzhalter

Aktuell unterstützt:

- `$CurrentUserGivenName$` -> `givenName`
- `$CurrentUserSurname$` -> `sn`
- `$CurrentUserTitle$` -> `title`
- `$CurrentUserTelephone$` -> `telephoneNumber`
- `$CurrentUserMail$` -> `mail`

Neue Variablen werden über `VariableMappings` ergänzt. Weitere Codeänderungen sind dafür in der Regel nicht nötig.

## Bereinigung leerer Felder

Wenn ein Attribut leer ist, wird der Platzhalter durch einen leeren String ersetzt. Anschließend kann optional eine Zeilenbereinigung greifen. Standardmäßig entfernt die Lösung:

- komplett leere Absätze
- Zeilen, die nur `Telefon:` enthalten
- Zeilen, die nur `E-Mail:` enthalten

Die Bereinigung ist konfigurierbar und kann pro Vorlage später erweitert werden.

## Lokaler Testbetrieb auf einem Einzelplatz

1. Vorlage nach `Templates\DK_Standard.docx` legen.
2. Konfiguration in [Config/signature-config.json](../Config/signature-config.json) prüfen.
3. Optional `VariableOverrides` für Offline-Tests befüllen.
4. PowerShell 5.1 im Repository öffnen.
5. Ausführen:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1
```

Optional ohne Ändern der Outlook-Standardwerte:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -SkipDefaultSignatureUpdate
```

Erzwungene Neuerzeugung:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -ForceRefresh
```

## Rollout über GPO

Das Wrapper-Skript [Scripts/Invoke-OutlookSignatureRollout.ps1](../Scripts/Invoke-OutlookSignatureRollout.ps1) ist für Login-Skript-Szenarien vorgesehen.

Beispiel:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "\\fileserver\Signatury\Scripts\Invoke-OutlookSignatureRollout.ps1" -ConfigPath "\\fileserver\Signatury\Config\signature-config.json"
```

Empfehlung für GPO:

- Dateien zentral unter `\\fileserver\Signatury` ablegen
- Clients nur lesend auf Vorlagen und Skripte zugreifen lassen
- Skript im Benutzerkontext ausführen

## Rollout über geplante Aufgabe

Beispiel-Aufruf:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\ProgramData\Signatury\Scripts\Invoke-OutlookSignatureRollout.ps1"
```

Empfohlene Parameter der Aufgabe:

- Ausführen bei Benutzeranmeldung
- Ausführen im Benutzerkontext
- Nur ausführen, wenn Netzwerk verfügbar ist, falls zentrale Vorlagen direkt benötigt werden

## Was die Lösung intern tut

1. Konfiguration laden und Pfade auflösen
2. Lokale Arbeits-, Cache-, Status- und Logordner anlegen
3. AD-Werte des aktuellen Benutzers lesen
4. Vorlage in den Cache spiegeln oder aus Cache verwenden
5. Fingerprint aus Vorlagenhash, Variablen und Exportoptionen berechnen
6. Nur bei Änderungen neu generieren
7. DOCX in Word öffnen, Platzhalter ersetzen und leere Zeilen bereinigen
8. Signatur als `.htm`, `.txt` und optional `.rtf` exportieren
9. Outlook-Standardwerte je nach Konfiguration in `Common\MailSettings` und optional im Standardprofil setzen

## Logging

Pro Lauf wird eine Logdatei geschrieben, lokal standardmäßig nach:

```text
.\Logs
```

Produktiv empfohlen:

```text
%LOCALAPPDATA%\Signatury\Logs
```

Im Log stehen unter anderem:

- Start und Ende
- erkannter Benutzer
- geladene Vorlage
- AD-Werte
- ersetzte Variablen
- erzeugte Dateien
- gesetzte Outlook-Defaults
- Fehler inklusive Stacktrace

## Troubleshooting

### Word COM kann nicht erstellt werden

- Prüfen, ob Microsoft Word installiert ist
- Word einmal interaktiv starten
- Prüfen, ob der Benutzer Word starten darf

### Active Directory ist nicht erreichbar

- Netzwerk- oder VPN-Verbindung prüfen
- Domänenanmeldung prüfen
- Für Einzelplatztests optional `VariableOverrides` setzen

### Vorlage wird nicht gefunden

- `TemplateRootPath` und `TemplateFileName` prüfen
- Bei zentraler Ablage Freigabe- und NTFS-Rechte prüfen
- Prüfen, ob bereits eine Cache-Kopie vorhanden ist

### Signatur wird erzeugt, aber nicht in Outlook als Standard gesetzt

- Sicherstellen, dass Classic Outlook verwendet wird
- Profilpfad unter `HKCU:\Software\Microsoft\Office\<Version>\Outlook` prüfen
- Outlook nach Änderung einmal neu starten
- Prüfen, ob der Benutzer mehrere Profile hat und welches Profil `DefaultProfile` ist
- Falls der Outlook-Signaturdialog danach nicht mehr bearbeitbar ist, `WriteDefaultsToProfileAccounts` testweise auf `false` setzen
- Für einen erzwungenen Neuaufbau kann zusätzlich `ResetManagedStateBeforeDeploy` auf `true` gesetzt werden

### Bilder fehlen in der HTML-Signatur

- Vorlage direkt in Word öffnen und Bild-Einbettung prüfen
- Prüfen, ob Word beim Export den Ordner `<Signaturname>_files` oder eine lokalisierte Variante erzeugt hat
- Sicherstellen, dass die Signatur in Outlook als HTML verwendet wird

### RTF wurde nicht erzeugt

- Der RTF-Export ist implementiert, wird aber bewusst nur als Best-Effort behandelt
- Wenn Word den RTF-Export für die konkrete Vorlage nicht sauber schafft, wird dies geloggt und die HTML/TXT-Ausgabe bleibt bestehen

## Einzeldateien

- Einstieg: [Scripts/Deploy-OutlookSignature.ps1](../Scripts/Deploy-OutlookSignature.ps1)
- Rollout-Wrapper: [Scripts/Invoke-OutlookSignatureRollout.ps1](../Scripts/Invoke-OutlookSignatureRollout.ps1)
- Kernlogik: [Scripts/OutlookSignature.Core.psm1](../Scripts/OutlookSignature.Core.psm1)
- Admin-Hinweise: [Administrator-Guide.md](Administrator-Guide.md)
- Selbsttests: [Tests/Invoke-OutlookSignatureSelfTest.ps1](../Tests/Invoke-OutlookSignatureSelfTest.ps1)
- Vorlagenhinweise: [Templates/README.md](../Templates/README.md)
