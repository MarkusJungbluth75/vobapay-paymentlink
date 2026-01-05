/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import QRCode from "qrcode";

const CONFIG_KEY = "vobapay_baseurl";
const HEADING_KEY = "vobapay_heading";
const DEFAULT_HEADING = "Flexibel digital bezahlen, mit dem vobapay QR Code";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Event listeners
    document.getElementById("saveConfig").onclick = saveConfiguration;
    document.getElementById("insertQR").onclick = insertPaymentQRCode;
    
    // Load saved configuration
    loadConfiguration();
  }
});

/**
 * Speichere die Base URL für den Paymentlink
 */
async function saveConfiguration() {
  const baseUrl = (document.getElementById("baseUrl") as HTMLInputElement).value;
  const heading = (document.getElementById("heading") as HTMLInputElement).value || DEFAULT_HEADING;
  const statusDiv = document.getElementById("config-status");
  
  if (!baseUrl || !isValidUrl(baseUrl)) {
    statusDiv.textContent = "❌ Ungültige URL!";
    statusDiv.style.color = "red";
    return;
  }
  
  // Store in localStorage
  localStorage.setItem(CONFIG_KEY, baseUrl);
  localStorage.setItem(HEADING_KEY, heading);
  statusDiv.textContent = "✅ Konfiguration gespeichert!";
  statusDiv.style.color = "green";
}

/**
 * Lade die gespeicherte Konfiguration
 */
function loadConfiguration() {
  const baseUrl = localStorage.getItem(CONFIG_KEY);
  const heading = localStorage.getItem(HEADING_KEY) || DEFAULT_HEADING;
  
  if (baseUrl) {
    (document.getElementById("baseUrl") as HTMLInputElement).value = baseUrl;
  }
  
  (document.getElementById("heading") as HTMLInputElement).value = heading;
  
  if (baseUrl) {
    const statusDiv = document.getElementById("config-status");
    statusDiv.textContent = "✅ Gespeicherte Konfiguration geladen";
    statusDiv.style.color = "green";
  }
}

/**
 * Validiere eine URL
 */
function isValidUrl(url: string): boolean {
  try {
    new URL(url);
    return true;
  } catch {
    return false;
  }
}

/**
 * Erstelle einen Paymentlink mit den angegebenen Parametern
 */
function createPaymentLink(baseUrl: string, amount: string, purpose: string): string {
  // Bereinige die Parameter
  const cleanAmount = parseFloat(amount).toFixed(2);
  const cleanPurpose = encodeURIComponent(purpose);
  
  // Erstelle den Link mit Parametern: a=Betrag, i=Verwendungszweck
  return `${baseUrl}?a=${cleanAmount}&i=${cleanPurpose}`;
}

/**
 * Extrahiere Zahlbetrag und Verwendungszweck aus dem Word-Dokument
 * Sucht nach Mustern wie:
 * - "Betrag: 123.45" oder "Betrag: 123,45 €"
 * - "Verwendungszweck: Text" oder "Zweck: Text"
 */
async function extractPaymentDataFromDocument(): Promise<{ amount: string; purpose: string }> {
  return await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const text = body.text;
    console.log("Document text:", text);

    // Suche nach Betrag (verschiedene Muster)
    let amount = "";
    const amountPatterns = [
      /Betrag:\s*([0-9]+[.,][0-9]{2})/i,
      /Betrag:\s*([0-9]+)/i,
      /Summe:\s*([0-9]+[.,][0-9]{2})/i,
      /Summe:\s*([0-9]+)/i,
      /([0-9]+[.,][0-9]{2})\s*€/,
      /€\s*([0-9]+[.,][0-9]{2})/
    ];

    for (const pattern of amountPatterns) {
      const match = text.match(pattern);
      if (match) {
        amount = match[1].replace(",", ".");
        console.log("Found amount:", amount);
        break;
      }
    }

    // Suche nach Verwendungszweck
    let purpose = "";
    const purposePatterns = [
      /Verwendungszweck:\s*([^\n]+)/i,
      /Zweck:\s*([^\n]+)/i,
      /Betreff:\s*([^\n]+)/i,
      /Referenz:\s*([^\n]+)/i
    ];

    for (const pattern of purposePatterns) {
      const match = text.match(pattern);
      if (match) {
        purpose = match[1].trim();
        console.log("Found purpose:", purpose);
        break;
      }
    }

    return { amount, purpose };
  });
}

/**
 * Konvertiere Blob zu Base64
 */
function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      const base64 = dataUrl.split(",")[1];
      resolve(base64);
    };
    reader.onerror = () => reject(new Error("Failed to read blob"));
    reader.readAsDataURL(blob);
  });
}

function sanitizeBase64(b64: string): string {
  return b64.replace(/\s+/g, "");
}

function isPngBase64(base64: string): boolean {
  try {
    const clean = sanitizeBase64(base64);
    const bin = atob(clean.substring(0, 64)); // decode a prefix
    // check for PNG signature (0x89 0x50 0x4E 0x47)
    return bin.charCodeAt(0) === 0x89 && bin.charCodeAt(1) === 0x50 && bin.charCodeAt(2) === 0x4e && bin.charCodeAt(3) === 0x47;
  } catch (e) {
    return false;
  }
}

/**
 * Füge Logo in die Mitte des QR-Codes ein
 */
async function embedLogoInQRCode(qrDataUrl: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    
    const qrImg = new Image();
    qrImg.onload = () => {
      canvas.width = qrImg.width;
      canvas.height = qrImg.height;
      
      // Zeichne QR-Code
      ctx.drawImage(qrImg, 0, 0);
      
      // Lade Logo
      const logoImg = new Image();
      logoImg.onload = () => {
        // Logo in der Mitte, ca. 20% der QR-Code-Größe
        const logoSize = Math.floor(canvas.width * 0.2);
        const logoX = (canvas.width - logoSize) / 2;
        const logoY = (canvas.height - logoSize) / 2;
        
        // Weißer Hintergrund für bessere Lesbarkeit
        ctx.fillStyle = "white";
        ctx.fillRect(logoX - 4, logoY - 4, logoSize + 8, logoSize + 8);
        
        // Zeichne Logo
        ctx.drawImage(logoImg, logoX, logoY, logoSize, logoSize);
        
        // Konvertiere zu Data URL
        resolve(canvas.toDataURL("image/png"));
      };
      logoImg.onerror = () => {
        console.warn("Logo konnte nicht geladen werden, verwende QR-Code ohne Logo");
        resolve(qrDataUrl);
      };
      logoImg.src = "assets/logo-filled.png";
    };
    qrImg.onerror = () => reject(new Error("Failed to load QR code image"));
    qrImg.src = qrDataUrl;
  });
}



/**
 * Füge einen Payment-QR-Code ins Word-Dokument ein
 */
export async function insertPaymentQRCode() {
  const baseUrl = localStorage.getItem(CONFIG_KEY);
  const errorDiv = document.getElementById("error-message");
  
  // Validierung
  if (!baseUrl) {
    errorDiv.textContent = "❌ Bitte speichern Sie zuerst die Konfiguration!";
    errorDiv.style.color = "red";
    return;
  }

  // Versuche Daten aus dem Dokument zu extrahieren
  let amount = (document.getElementById("amount") as HTMLInputElement).value;
  let purpose = (document.getElementById("purpose") as HTMLInputElement).value;

  // Wenn Felder leer sind, automatisch aus Dokument auslesen
  if (!amount || !purpose) {
    errorDiv.textContent = "🔍 Suche Daten im Dokument...";
    errorDiv.style.color = "blue";
    
    try {
      const extracted = await extractPaymentDataFromDocument();
      if (!amount && extracted.amount) {
        amount = extracted.amount;
        (document.getElementById("amount") as HTMLInputElement).value = amount;
      }
      if (!purpose && extracted.purpose) {
        purpose = extracted.purpose;
        (document.getElementById("purpose") as HTMLInputElement).value = purpose;
      }
    } catch (extractErr) {
      console.warn("Failed to extract data from document:", extractErr);
    }
  }
  
  if (!amount || parseFloat(amount) <= 0) {
    errorDiv.textContent = "❌ Kein gültiger Betrag gefunden! Bitte manuell eingeben.";
    errorDiv.style.color = "red";
    return;
  }
  
  if (!purpose || purpose.trim() === "") {
    errorDiv.textContent = "❌ Kein Verwendungszweck gefunden! Bitte manuell eingeben.";
    errorDiv.style.color = "red";
    return;
  }
  
  try {
    // Erstelle den Paymentlink
    const paymentLink = createPaymentLink(baseUrl, amount, purpose);

    // Versuche zuerst, mit der eingebundenen qrcode-Bibliothek eine Data-URL zu erzeugen
    let imageBase64: string | null = null;
    try {
      const qrCodeDataUrl = await QRCode.toDataURL(paymentLink, {
        width: 400,
        margin: 2,
        errorCorrectionLevel: "H", // Hohe Fehlerkorrektur für Logo-Einbettung
        color: { dark: "#000000", light: "#FFFFFF" },
      });
      
      // Füge Logo in QR-Code ein
      const qrWithLogo = await embedLogoInQRCode(qrCodeDataUrl);
      imageBase64 = qrWithLogo.split(",")[1];
      console.log("Local QRCode with logo generated successfully");
    } catch (genErr) {
      console.warn("QRCode.toDataURL failed, falling back to remote QR service:", genErr);
    }

    // Fallback: lade PNG von api.qrserver.com und konvertiere zu Base64
    if (!imageBase64) {
      try {
        const qrServerUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(
          paymentLink
        )}`;
        console.log("Fetching QR from server:", qrServerUrl);
        const resp = await fetch(qrServerUrl);
        if (!resp.ok) throw new Error(`QR server returned ${resp.status}`);
        
        const blob = await resp.blob();
        console.log("Blob received, size:", blob.size, "type:", blob.type);
        
        imageBase64 = await blobToBase64(blob);
        console.log("Base64 from remote QR: OK, length=" + imageBase64.length);
      } catch (fetchErr) {
        console.error("Failed to fetch QR image from server:", fetchErr);
        throw fetchErr;
      }
    }

    if (!imageBase64 || imageBase64.length === 0) {
      throw new Error("Could not generate QR code image");
    }

    // Füge den QR-Code ins Word-Dokument ein
    try {
      await Word.run(async (context) => {
      // Hole die konfigurierte Überschrift
      const heading = localStorage.getItem(HEADING_KEY) || DEFAULT_HEADING;
      
      // Erstelle Überschrift
      const headingParagraph = context.document.body.insertParagraph(heading, Word.InsertLocation.end);
      headingParagraph.alignment = Word.Alignment.center;
      headingParagraph.font.size = 14;
      headingParagraph.font.bold = true;
      headingParagraph.spaceAfter = 10;
      
      // Erstelle einen neuen Absatz für den QR-Code
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph.alignment = Word.Alignment.center;

      try {
        // Bereinige Base64 (entferne Zeilenumbrüche) und validiere PNG-Header
        imageBase64 = sanitizeBase64(imageBase64 as string);
        if (!isPngBase64(imageBase64)) {
          console.warn("Base64 does not look like PNG; length=", imageBase64.length);
          throw new Error("Generated image is not a valid PNG (invalid header)");
        }

        // Füge das Bild ein (100x100 Pixel = ca. 3,5cm)
        const inlineShape = paragraph.getRange().insertInlinePictureFromBase64(imageBase64, Word.InsertLocation.start);
        inlineShape.width = 100;
        inlineShape.height = 100;
        
        // Füge Hyperlink zum Bild hinzu
        inlineShape.hyperlink = paymentLink;
        console.log("QR image with hyperlink inserted successfully");
      } catch (insertErr) {
        console.error("Image insertion error:", insertErr);
        // Wenn OfficeExtension.Error.debugInfo vorhanden ist, logge detaillierte Infos
        if ((insertErr as any).debugInfo) {
          console.error("OfficeExtension.Error.debugInfo:", JSON.stringify((insertErr as any).debugInfo));
        }
        throw new Error(`Failed to insert image: ${(insertErr as any).message || insertErr}`);
      }

      // Füge Informationen unter dem QR-Code hinzu
      const infoLine = context.document.body.insertParagraph("", Word.InsertLocation.end);
      infoLine.alignment = Word.Alignment.center;
      infoLine.font.size = 10;
      infoLine.font.italic = true;
      infoLine.insertText(`Betrag: ${amount}€ | Zweck: ${purpose}`, Word.InsertLocation.end);

      // Speichern erzwingen
      await context.sync();
      console.log("Word.run completed successfully");
      });
    } catch (wordRunErr) {
      console.error("Word.run error:", wordRunErr);
      if ((wordRunErr as any).debugInfo) {
        console.error("OfficeExtension.Error.debugInfo:", JSON.stringify((wordRunErr as any).debugInfo));
        errorDiv.textContent = `❌ Fehler (Office): ${(wordRunErr as any).message}. Debug: ${JSON.stringify((wordRunErr as any).debugInfo)}`;
      } else {
        errorDiv.textContent = `❌ Fehler: ${(wordRunErr as any).message || wordRunErr}`;
      }
      errorDiv.style.color = "red";
      return;
    }

    // Erfolgs-Nachricht
    errorDiv.textContent = "✅ QR-Code erfolgreich eingefügt!";
    errorDiv.style.color = "green";

    // Eingabefelder leeren
    (document.getElementById("amount") as HTMLInputElement).value = "";
    (document.getElementById("purpose") as HTMLInputElement).value = "";
  } catch (error) {
    console.error("Fehler beim Einfügen des QR-Codes:", error);
    errorDiv.textContent = `❌ Fehler: ${error && (error as any).message ? (error as any).message : error}`;
    errorDiv.style.color = "red";
  }
}
