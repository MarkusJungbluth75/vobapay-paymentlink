# VobaPay Payment Link QR-Code Generator – Installationsanleitung

## Überblick
Dieses Office Add-in generiert QR-Codes für VobaPay Zahlungslinks und fügt diese direkt in Word-Dokumente ein.

### Features
- ✅ Konfigurierbare Base URL für Zahlungslinks
- ✅ Dynamische Parameter: Zahlbetrag (`a`) und Verwendungszweck (`i`)
- ✅ QR-Code-Generierung und Einbindung in Word-Dokumente
- ✅ Lokale Konfigurationsspeicherung per localStorage

---

## Installation & Setup

### 1. Voraussetzungen
- Node.js 18, 20 oder 22
- Word für Windows (Beta Channel, Build 18514+) oder macOS
- Microsoft 365 Konto

### 2. Dependencies installieren
```bash
npm install
```

Die `qrcode`-Bibliothek ist bereits installiert für QR-Code-Generierung.

### 3. Dev Server starten
```bash
npm run dev-server
```
Der Server läuft auf `https://localhost:3000` (HTTPS mit selbstsigniertem Zertifikat).

### 4. Add-in debuggen in Word
```bash
npm run start:desktop:word
```

---

## Verwendung

### Schritt 1: Base URL konfigurieren
1. Öffne die **Konfiguration**-Sektion im Add-in
2. Gib deine VobaPay-Base-URL ein, z.B.:
   ```
   https://payment.vobapay.com/pay
   ```
3. Klick **„Konfiguration speichern"** (wird lokal gespeichert)

### Schritt 2: Payment Link einfügen
1. Gib den **Zahlbetrag** ein (z.B. `99.99`)
2. Gib einen **Verwendungszweck** ein (z.B. `Rechnung #12345`)
3. Klick **„QR-Code einfügen"**

### Ergebnis
Der QR-Code wird ins Dokument eingefügt mit:
- **QR-Code** (300×300px, zentriert)
- **Info-Zeile** darunter mit Betrag und Zweck

---

## Paymentlink-Format

Der generierte Link hat folgende Struktur:
```
https://payment.vobapay.com/pay?a=99.99&i=Rechnung%20%2312345
```

**Parameter:**
- `a` = Zahlbetrag in Euro (mit Dezimalpunkt)
- `i` = Verwendungszweck (URL-kodiert)

---

## Technische Details

### Dateien
- `src/taskpane/taskpane.html` – UI mit Formularen
- `src/taskpane/word.ts` – Logik für Word (QR-Code-Generierung)
- `src/taskpane/taskpane.css` – Styling
- `appPackage/manifest.json` – Add-in-Konfiguration

### Dependencies
- `qrcode` – QR-Code-Generierung zu PNG
- `office-js` – Word API

### Build & Deploy
```bash
# Development Build
npm run build:dev

# Production Build
npm run build

# Manifest validieren
npm run validate

# Production sideload
npx office-addin-dev-settings sideload ./dist/manifest.json
```

---

## Fehlerbehebung

### Problem: „Ungültige URL"
- Stelle sicher, dass die URL mit `https://` oder `http://` beginnt
- Beispiel: ✅ `https://payment.vobapay.com/pay` ❌ `payment.vobapay.com/pay`

### Problem: „QR-Code wird nicht eingefügt"
- Überprüfe die Browser-Konsole (F12) auf Fehler
- Stelle sicher, dass die Basis-Konfiguration gespeichert ist
- Alle Felder müssen gefüllt sein (Betrag, Verwendungszweck)

### Problem: Dev-Server lädt nicht
```bash
# Cache leeren und neu starten
npm run build:dev
npm run dev-server
```

---

## Production Deployment

1. **Endpoint URL aktualisieren:**
   - Öffne `webpack.config.js`
   - Ändere `urlProd` auf deine Production URL
   
   Oder setze in `env/.env.dev`:
   ```
   ADDIN_ENDPOINT=https://your-domain.com/
   ```

2. **Production Build:**
   ```bash
   npm run build
   ```

3. **Manifest aktualisieren & sideload:**
   ```bash
   npx office-addin-dev-settings sideload ./dist/manifest.json
   ```

---

## Support
Bei Fragen oder Problemen: siehe Logs im Browser (F12) oder Terminal-Output.
