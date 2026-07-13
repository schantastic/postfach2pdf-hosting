/* global Office, PDFLib, Blob, URL, Image, atob, btoa, window */
/**
 * Postfach2PDF Core: gemeinsame Hilfsfunktionen fuer Lese- und Verfassen-Taskpane.
 * Haengt sich als window.Postfach2PdfCore an; haengt selbst nur von PDFLib und
 * Office.js ab (kein DOM-Zugriff), damit beide Taskpanes es unveraendert
 * wiederverwenden koennen.
 */
(function (global) {
  "use strict";

  var MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024; // 20 MB Sicherheitslimit fuer diese Version
  var SUPPORTED_RASTER_IMAGE_TYPES = [
    "image/jpeg",
    "image/jpg",
    "image/png",
    "image/gif",
    "image/bmp",
    "image/webp",
  ];
  var PAGE_SIZE_A4 = [595.28, 841.89]; // Punkt (pt)

  var WINDOWS_RESERVED_NAMES = [
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
  ];

  // ---------------------------------------------------------------------
  // Dateinamen (Windows-sicher)
  // ---------------------------------------------------------------------

  function isPrintableChar(ch) {
    var code = ch.charCodeAt(0);
    return code >= 32 && code !== 127;
  }

  function sanitizeFileNamePart(value) {
    if (!value) {
      return "";
    }
    var withoutControlChars = value.split("").filter(isPrintableChar).join("");
    var withoutForbiddenChars = withoutControlChars.replace(/[<>:"/\\|?*]/g, "");
    return withoutForbiddenChars.replace(/\s+/g, " ").trim();
  }

  function sanitizeFileNameStem(stem) {
    var cleaned = stem.replace(/[. ]+$/g, "");
    if (cleaned === "") {
      cleaned = "unbenannt";
    }
    if (WINDOWS_RESERVED_NAMES.indexOf(cleaned.toUpperCase()) !== -1) {
      cleaned = "_" + cleaned;
    }
    return cleaned.slice(0, 150);
  }

  // ---------------------------------------------------------------------
  // Bytes / Base64
  // ---------------------------------------------------------------------

  function formatBytes(bytes) {
    if (typeof bytes !== "number" || isNaN(bytes)) {
      return "unbekannte Groesse";
    }
    if (bytes < 1024) {
      return bytes + " B";
    }
    if (bytes < 1024 * 1024) {
      return (bytes / 1024).toFixed(1) + " KB";
    }
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }

  function base64ToUint8Array(base64) {
    var binaryString = atob(base64);
    var bytes = new Uint8Array(binaryString.length);
    for (var i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  }

  function uint8ArrayToBase64(bytes) {
    var chunkSize = 0x8000;
    var chunks = [];
    for (var i = 0; i < bytes.length; i += chunkSize) {
      var chunk = bytes.subarray(i, i + chunkSize);
      chunks.push(String.fromCharCode.apply(null, chunk));
    }
    return btoa(chunks.join(""));
  }

  function sanitizeForStandardFont(text) {
    // Standard-PDF-Fonts (WinAnsi) decken nur Latin-1 (Codepunkte 0-255) ab.
    return String(text).replace(/[^\x00-\xFF]/g, "?");
  }

  function wrapText(text, maxCharsPerLine) {
    var words = String(text).split(" ");
    var lines = [];
    var current = "";
    words.forEach(function (word) {
      var candidate = current ? current + " " + word : word;
      if (candidate.length > maxCharsPerLine && current) {
        lines.push(current);
        current = word;
      } else {
        current = candidate;
      }
    });
    if (current) {
      lines.push(current);
    }
    return lines;
  }

  // ---------------------------------------------------------------------
  // Office.js-Wrapper (Read + Compose gemeinsam nutzbar)
  // ---------------------------------------------------------------------

  function getAttachmentContent(item, attachmentId) {
    return new Promise(function (resolve, reject) {
      item.getAttachmentContentAsync(attachmentId, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function isConvertibleAttachment(attachment) {
    var contentType = (attachment.contentType || "").toLowerCase();
    if (contentType === "application/pdf" || /\.pdf$/i.test(attachment.name || "")) {
      return true;
    }
    return SUPPORTED_RASTER_IMAGE_TYPES.indexOf(contentType) !== -1;
  }

  // ---------------------------------------------------------------------
  // Bilder rasterisieren
  // ---------------------------------------------------------------------

  function loadImageElement(blob) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () {
        resolve({ img: img, url: url });
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Bild konnte nicht dekodiert werden"));
      };
      img.src = url;
    });
  }

  function rasterizeToPngBytes(bytes, mimeType) {
    var blob = new Blob([bytes], { type: mimeType });
    return loadImageElement(blob).then(function (loaded) {
      try {
        var canvas = document.createElement("canvas");
        canvas.width = loaded.img.naturalWidth;
        canvas.height = loaded.img.naturalHeight;
        var ctx = canvas.getContext("2d");
        ctx.drawImage(loaded.img, 0, 0);
        var dataUrl = canvas.toDataURL("image/png");
        var base64 = dataUrl.split(",")[1];
        return {
          bytes: base64ToUint8Array(base64),
          width: canvas.width,
          height: canvas.height,
        };
      } finally {
        URL.revokeObjectURL(loaded.url);
      }
    });
  }

  // ---------------------------------------------------------------------
  // PDF-Seiten erzeugen
  // ---------------------------------------------------------------------

  async function addPlaceholderPage(mergedPdf, attachment, reason) {
    var page = mergedPdf.addPage(PAGE_SIZE_A4);
    var font = await mergedPdf.embedFont(PDFLib.StandardFonts.Helvetica);
    var boldFont = await mergedPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
    var margin = 50;
    var y = PAGE_SIZE_A4[1] - margin;

    function drawLine(text, opts) {
      var options = opts || {};
      var size = options.size || 11;
      page.drawText(sanitizeForStandardFont(text), {
        x: margin,
        y: y,
        size: size,
        font: options.font || font,
        color: PDFLib.rgb(0, 0, 0),
      });
      y -= size + 8;
    }

    drawLine("Anhang nicht eingebettet", { font: boldFont, size: 16 });
    y -= 4;
    drawLine("Dateiname: " + (attachment.name || "unbekannt"));
    drawLine("Typ: " + (attachment.contentType || "unbekannt"));
    drawLine("Groesse: " + formatBytes(attachment.size));
    y -= 8;
    wrapText(reason, 85).forEach(function (line) {
      drawLine(line);
    });

    return { status: "placeholder", reason: reason };
  }

  async function embedPdfAttachment(mergedPdf, bytes, attachment) {
    try {
      var srcDoc = await PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
      var copiedPages = await mergedPdf.copyPages(srcDoc, srcDoc.getPageIndices());
      copiedPages.forEach(function (p) {
        mergedPdf.addPage(p);
      });
      return { status: "embedded" };
    } catch (error) {
      var reason =
        "PDF-Anhang konnte nicht eingebettet werden (evtl. verschluesselt oder beschaedigt): " +
        (error && error.message ? error.message : error);
      return addPlaceholderPage(mergedPdf, attachment, reason);
    }
  }

  async function embedImageAttachment(mergedPdf, bytes, contentType, attachment) {
    try {
      var rasterized = await rasterizeToPngBytes(bytes, contentType);
      var pngImage = await mergedPdf.embedPng(rasterized.bytes);
      var margin = 36;
      var maxW = PAGE_SIZE_A4[0] - margin * 2;
      var maxH = PAGE_SIZE_A4[1] - margin * 2;
      var scale = Math.min(maxW / rasterized.width, maxH / rasterized.height);
      var drawWidth = rasterized.width * scale;
      var drawHeight = rasterized.height * scale;
      var page = mergedPdf.addPage(PAGE_SIZE_A4);
      page.drawImage(pngImage, {
        x: (PAGE_SIZE_A4[0] - drawWidth) / 2,
        y: (PAGE_SIZE_A4[1] - drawHeight) / 2,
        width: drawWidth,
        height: drawHeight,
      });
      return { status: "embedded" };
    } catch (error) {
      var reason = "Bild-Anhang konnte nicht eingebettet werden: " + (error && error.message ? error.message : error);
      return addPlaceholderPage(mergedPdf, attachment, reason);
    }
  }

  // item: das Office.js-Mailbox-Item (Read- oder Compose-Modus), das
  // getAttachmentContentAsync bereitstellt.
  async function embedAttachmentIntoPdf(item, mergedPdf, attachment) {
    try {
      if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.Cloud) {
        return addPlaceholderPage(
          mergedPdf,
          attachment,
          "Cloud-Anhang (Freigabelink): Inhalt kann in dieser Version nicht automatisch gelesen werden."
        );
      }
      if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.Item) {
        return addPlaceholderPage(
          mergedPdf,
          attachment,
          "Eingebettetes Outlook-Element (z. B. Mail/Termin/Kontakt) wird in dieser Version nicht automatisch konvertiert."
        );
      }
      if (typeof attachment.size === "number" && attachment.size > MAX_ATTACHMENT_BYTES) {
        return addPlaceholderPage(
          mergedPdf,
          attachment,
          "Anhang wurde uebersprungen: Groesse (" +
            formatBytes(attachment.size) +
            ") liegt ueber dem Limit von " +
            formatBytes(MAX_ATTACHMENT_BYTES) +
            " fuer diese Version."
        );
      }

      var content = await getAttachmentContent(item, attachment.id);

      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        return addPlaceholderPage(
          mergedPdf,
          attachment,
          "Anhangsinhalt liegt in einem Format vor (" + content.format + "), das in dieser Version nicht verarbeitet wird."
        );
      }

      var bytes = base64ToUint8Array(content.content);
      var contentType = (attachment.contentType || "").toLowerCase();

      if (contentType === "application/pdf" || /\.pdf$/i.test(attachment.name || "")) {
        return embedPdfAttachment(mergedPdf, bytes, attachment);
      }

      if (SUPPORTED_RASTER_IMAGE_TYPES.indexOf(contentType) !== -1) {
        return embedImageAttachment(mergedPdf, bytes, contentType, attachment);
      }

      return addPlaceholderPage(
        mergedPdf,
        attachment,
        "Automatische Konvertierung dieses Dateityps (" +
          (attachment.contentType || "unbekannt") +
          ") ist in dieser Version nicht implementiert."
      );
    } catch (error) {
      console.error("Postfach2PDF: Anhang konnte nicht verarbeitet werden", attachment.name, error);
      return addPlaceholderPage(
        mergedPdf,
        attachment,
        "Anhang konnte nicht gelesen werden: " + (error && error.message ? error.message : error)
      );
    }
  }

  async function addPageNumbers(mergedPdf) {
    var font = await mergedPdf.embedFont(PDFLib.StandardFonts.Helvetica);
    var pages = mergedPdf.getPages();
    var total = pages.length;
    var margin = 36;

    pages.forEach(function (page, index) {
      var width = page.getWidth();
      var text = "Seite " + (index + 1) + " von " + total;
      var size = 8;
      var textWidth = font.widthOfTextAtSize(text, size);
      page.drawLine({
        start: { x: margin, y: margin - 6 },
        end: { x: width - margin, y: margin - 6 },
        thickness: 0.5,
        color: PDFLib.rgb(0.8, 0.8, 0.8),
      });
      page.drawText(text, {
        x: (width - textWidth) / 2,
        y: margin - 18,
        size: size,
        font: font,
        color: PDFLib.rgb(0.4, 0.4, 0.4),
      });
      page.drawText("Postfach2PDF", {
        x: margin,
        y: margin - 18,
        size: size,
        font: font,
        color: PDFLib.rgb(0.4, 0.4, 0.4),
      });
    });
  }

  function downloadBytes(bytes, fileName) {
    var blob = new Blob([bytes], { type: "application/pdf" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  global.Postfach2PdfCore = {
    MAX_ATTACHMENT_BYTES: MAX_ATTACHMENT_BYTES,
    SUPPORTED_RASTER_IMAGE_TYPES: SUPPORTED_RASTER_IMAGE_TYPES,
    PAGE_SIZE_A4: PAGE_SIZE_A4,
    sanitizeFileNamePart: sanitizeFileNamePart,
    sanitizeFileNameStem: sanitizeFileNameStem,
    formatBytes: formatBytes,
    base64ToUint8Array: base64ToUint8Array,
    uint8ArrayToBase64: uint8ArrayToBase64,
    sanitizeForStandardFont: sanitizeForStandardFont,
    wrapText: wrapText,
    getAttachmentContent: getAttachmentContent,
    isConvertibleAttachment: isConvertibleAttachment,
    rasterizeToPngBytes: rasterizeToPngBytes,
    addPlaceholderPage: addPlaceholderPage,
    embedAttachmentIntoPdf: embedAttachmentIntoPdf,
    addPageNumbers: addPageNumbers,
    downloadBytes: downloadBytes,
  };
})(window);
