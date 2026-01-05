# 🚀 Word Add-in testen – Schnellstart

## Problem
Das automatische Sideloading schlägt aktuell auf Grund von M365 Manifest-Validierungsproblemen fehl.

## Lösungen

### **Lösung 1: Test-Seite (schnellster Weg)**
Öffne lokal die Test-HTML, um die QR-Code-Generierung zu testen:

```bash
# Dev-Server läuft bereits auf localhost:3000
open https://localhost:3000/test.html
```

Dort kannst du:
- ✅ Base URL speichern
- ✅ Zahlbetrag & Zweck eingeben
- ✅ QR-Code generieren und sehen

---

### **Lösung 2: Manuelles Laden in Word (für echten Test)**

**Für macOS/Windows:**

1. **Word öffnen** → Neues Dokument
2. **Insert-Tab** → **Get Add-ins** (oder **Einfügen** → **Add-Ins abrufen**)
3. Wähle **"My Add-ins"** (Meine Add-Ins) → **"Upload My Add-in"** (Mein Add-In hochladen)
4. Navigiere zu:
   ```
   /Users/markusjungbluth/AgentsToolkitProjects/vobapay_paymentlink/dist/manifest.json
   ```
5. Klick **Upload**

**Erwartetes Ergebnis:**
- Neue Button-Gruppe **"VobaPay"** auf dem **Home-Reiter**
- Button **"QR-Code"**
- Klick öffnet das Taskpane auf der rechten Seite

---

### **Lösung 3: Debugging & Troubleshooting**

Falls der Button nicht sichtbar ist:

```bash
# Stelle sicher, dass der Dev-Server läuft:
npm run dev-server

# Build erneuern:
npm run build:dev

# Browser-Cache leeren (Strg+Shift+Delete oder Cmd+Shift+Delete)

# In Word F12 drücken und die Konsole prüfen
```

---

## Datei-Struktur

```
src/taskpane/
├── taskpane.html     ← UI mit Formularen
├── taskpane.ts       ← Event-Listener & Logik
├── word.ts           ← Word.run() & QR-Code Insert
└── taskpane.css      ← Styling

dist/
├── manifest.json     ← Zum Laden in Word
├── taskpane.html     ← Kompilierte HTML
└── taskpane.js       ← Kompilierte TypeScript
```

---

## Schnell-Checkliste

- [ ] Dev-Server läuft: `npm run build:dev && npm run dev-server`
- [ ] Manifest ist gültig: `npm run validate`
- [ ] Test-Seite funktioniert: https://localhost:3000/test.html
- [ ] Add-in in Word geladen (Button sichtbar)
- [ ] QR-Code wird bei Button-Klick ins Dokument eingefügt

---

Wenn du weitere Probleme hast, schreib mir Bescheid! 🎯
