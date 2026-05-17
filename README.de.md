# VSKeyExtractor

VSKeyExtractor ist ein kleines Kommandozeilenwerkzeug, das versucht, den Lizenzschlüssel aus einer lokalen Visual Studio-Installation zu extrahieren.

Es durchsucht bekannte Visual Studio-Produkt-Einträge, liest die verschlüsselten Lizenzdaten aus der Windows-Registrierung, entschlüsselt sie im Kontext des aktuellen Benutzers und sucht nach einem Schlüssel im Format `AAAAA-BBBBB-CCCCC-DDDDD-EEEEE`.

## Funktionen

- Liest Visual-Studio-Lizenzdaten aus der Windows-Registrierung
- Unterstützt mehrere Visual-Studio-Produktgenerationen
- Versucht, geschützte Lizenzdaten mit dem aktuellen Benutzerprofil zu entschlüsseln
- Gibt gefundene schlüsselartige Werte in der Konsole aus

## Voraussetzungen

- Windows
- Eine Visual-Studio-Installation mit vorhandenen lokalen Lizenzdaten
- Die benötigte .NET-Runtime für das jeweilige Projekt

## Funktionsweise

Das Tool prüft Registrierungswerte an den bekannten Visual-Studio-Lizenzpfaden und versucht anschließend, die gespeicherten Daten zu entschlüsseln. Wenn im entschlüsselten Text ein passender Schlüssel gefunden wird, wird er in die Standardausgabe geschrieben.

Falls kein Schlüssel gefunden wird, fährt das Tool einfach mit dem nächsten bekannten Produkt-Eintrag fort.

## Anwendung starten

Du kannst das Projekt direkt mit `dotnet run` starten.

### .NET Framework / Windows-Target

```powershell
dotnet run --project .\VSKeyExtractor.csproj
```

### SDK-style-Projekt

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj
```

### Bestimmtes Target Framework ausführen

Wenn du das SDK-style-Projekt verwendest und ein bestimmtes Framework starten möchtest, kannst du Folgendes nutzen:

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj -f net8.0-windows
```

Weitere unterstützte Target Frameworks können `net9.0-windows` oder `net10.0-windows` sein, abhängig vom installierten .NET SDK.

## Beispielausgabe

```text
Found key for Visual Studio 2022 Professional: ABCDE-FGHIJ-KLMNO-PQRST-UVWXY
```

## Hinweise

- Führe das Tool auf dem Rechner aus, auf dem Visual Studio aktiviert wurde.
- Das Tool kann nur Daten lesen, die für das aktuelle Benutzerprofil verfügbar sind.
- Einige Einträge können nicht verfügbar, fehlerhaft oder geschützt sein; das ist erwartbar.

## Übersetzungen

- [English](README.md)
- [Français](README.fr.md)
- [Español](README.es.md)
