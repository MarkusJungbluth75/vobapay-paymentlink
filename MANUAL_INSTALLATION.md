# Manuelles Laden des VobaPay Payment QR-Code Add-ins in Word

Da das automatische Sideloading auf Manifest-Validierungsprobleme trifft, folge diesen Schritten zum manuellen Laden:

## Option 1: Über Datei-Dialog in Word

1. **Word öffnen** und ein neues Dokument erstellen
2. **"Insert"** (oder **Einfügen**) Menü → **Get Add-ins** (oder **Add-Ins abrufen**)
3. **"My Add-ins"** → **Upload My Add-in**
4. Wähle diese Datei aus:
   ```
   /Users/markusjungbluth/AgentsToolkitProjects/vobapay_paymentlink/dist/manifest.json
   ```
5. Klick **Upload**

## Option 2: Über den Developer Server (für lokale Entwicklung)

Stelle sicher, dass:
1. Der Dev-Server läuft:
   ```bash
   npm run dev-server
   ```
   (sollte auf `https://localhost:3000` laufen)

2. In Word:
   - **Insert** → **Get Add-ins** → **My Add-ins** → **Upload My Add-in**
   - Wähle: `/Users/markusjungbluth/AgentsToolkitProjects/vobapay_paymentlink/dist/manifest.json`

## Option 3: Troubleshooting bei Manifest-Fehlern

Falls der Fehler `"must NOT have additional properties"` erscheint:

1. Prüfe die Manifest auf ungültige Felder
2. Validiere lokal: `npm run validate`
3. Lade neu: Der Browser-Cache kann alte Manifeste zwischenspeichern

## Erwartetes Verhalten nach erfolgreichem Laden

- Ein neuer **"VobaPay"**-Button sollte auf dem **Home**-Reiter in Word erscheinen
- Button Label: **"QR-Code"**
- Klick öffnet einen **Taskpane** (Seitenleiste) mit:
  - Konfigurationssektion (Base URL eingeben)
  - Payment Link Section (Betrag + Zweck eingeben)

## Lokal debuggen

Wenn du auf Probleme stößt, öffne die Browser-Konsole:
- **Word Web**: F12 → Developer Tools
- **Word Desktop**: F12 könnte auch arbeiten oder prüfe die Logs im Terminal

Logs sollten im Terminal angezeigt werden, wo du `npm run dev-server` ausgeführt hast.
