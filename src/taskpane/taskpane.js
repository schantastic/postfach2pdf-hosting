/* global Office, PDFLib, html2pdf, Postfach2PdfCore, document, window */

(function () {
  "use strict";

  var el = {};
  var currentAttachments = [];
  var batchSelection = [];

  var SETTINGS_KEYS = {
    loadExternalImages: "mailpdf.loadExternalImages",
    includeHeaders: "mailpdf.includeHeaders",
    pageNumbers: "mailpdf.pageNumbers",
  };

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      return;
    }

    el.app = document.getElementById("mp-app");
    el.batch = document.getElementById("mp-batch");
    el.unsupported = document.getElementById("mp-unsupported");
    el.shared = document.getElementById("mp-shared");
    el.subject = document.getElementById("mp-subject");
    el.from = document.getElementById("mp-from");
    el.date = document.getElementById("mp-date");
    el.attachmentList = document.getElementById("mp-attachment-list");
    el.noAttachments = document.getElementById("mp-no-attachments");
    el.loadExternalImages = document.getElementById("mp-load-external-images");
    el.includeHeaders = document.getElementById("mp-include-headers");
    el.pageNumbers = document.getElementById("mp-page-numbers");
    el.generate = document.getElementById("mp-generate");
    el.status = document.getElementById("mp-status");
    el.renderRoot = document.getElementById("mp-render-root");
    el.result = document.getElementById("mp-result");
    el.resultList = document.getElementById("mp-result-list");

    el.batchList = document.getElementById("mp-batch-list");
    el.batchEmpty = document.getElementById("mp-batch-empty");
    el.batchMerge = document.getElementById("mp-batch-merge");
    el.batchIncludeHeaders = document.getElementById("mp-batch-include-headers");
    el.batchPageNumbers = document.getElementById("mp-batch-page-numbers");
    el.batchStart = document.getElementById("mp-batch-start");
    el.debug = document.getElementById("mp-debug");

    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      el.unsupported.hidden = false;
      el.unsupported.textContent =
        "Dieses Outlook unterstuetzt die fuer Postfach2PDF benoetigte Mailbox-API (1.8, fuer Anhangszugriff) nicht.";
      return;
    }

    el.shared.hidden = false;

    debugLog(
      "Office.onReady - currentItem vorhanden: " + !!currentItem() +
        " - Mailbox 1.13 unterstuetzt: " + Office.context.requirements.isSetSupported("Mailbox", "1.13")
    );

    if (currentItem()) {
      initSingleItemMode();
    } else if (Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
      initBatchMode();
    } else {
      el.unsupported.hidden = false;
      el.unsupported.textContent =
        "Kein geoeffnetes Element und keine Mehrfachauswahl-API (Mailbox 1.13) verfuegbar. " +
        "Bitte eine E-Mail einzeln oeffnen.";
    }
  });

  function currentItem() {
    return Office.context.mailbox.item;
  }

  function nonInlineAttachments(item) {
    return (item.attachments || []).filter(function (a) {
      return !a.isInline;
    });
  }

  function formatAddress(addr) {
    if (!addr) {
      return "";
    }
    if (addr.displayName && addr.emailAddress) {
      return addr.displayName + " <" + addr.emailAddress + ">";
    }
    return addr.displayName || addr.emailAddress || "";
  }

  function formatAddressList(list) {
    if (!list || list.length === 0) {
      return "";
    }
    return list.map(formatAddress).join("; ");
  }

  function formatDate(date) {
    if (!date) {
      return "";
    }
    try {
      return new Intl.DateTimeFormat("de-DE", {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(date);
    } catch (e) {
      return date.toString();
    }
  }

  // Temporaer zur Fehlersuche: schreibt zusaetzlich zu console.log direkt
  // sichtbar auf die Seite, weil die Browser-Konsole in Outlooks
  // verschachtelten iframes schwer zu finden ist.
  function debugLog(text) {
    console.log("Postfach2PDF: " + text);
    if (el.debug) {
      el.debug.textContent += new Date().toLocaleTimeString("de-DE") + "  " + text + "\n";
    }
  }

  function setStatus(text, state) {
    el.status.textContent = text;
    if (state) {
      el.status.setAttribute("data-state", state);
    } else {
      el.status.removeAttribute("data-state");
    }
  }

  function renderResults(results) {
    el.resultList.innerHTML = "";
    results.forEach(function (entry) {
      var li = document.createElement("li");
      var icon = document.createElement("span");
      icon.className = "mp-result-icon";
      var label = document.createElement("span");

      if (entry.status === "embedded") {
        icon.textContent = "✓";
        icon.setAttribute("data-kind", "ok");
        label.textContent = entry.name + " — " + (entry.detail || "eingebettet");
      } else if (entry.status === "skipped") {
        icon.textContent = "–";
        icon.setAttribute("data-kind", "skip");
        label.textContent = entry.name + " — " + (entry.detail || "uebersprungen (manuell abgewaehlt)");
      } else {
        icon.textContent = "⚠";
        icon.setAttribute("data-kind", "warn");
        label.textContent = entry.name + " — " + (entry.reason || entry.detail || "Fehler");
      }

      li.appendChild(icon);
      li.appendChild(label);
      el.resultList.appendChild(li);
    });
    el.result.hidden = results.length === 0;
  }

  // =======================================================================
  // Einzel-Modus (eine geoeffnete/ausgewaehlte E-Mail)
  // =======================================================================

  function initSingleItemMode() {
    el.app.hidden = false;
    renderSummary();
    renderAttachmentList();
    loadPersistedSettings();
    wireSettingsPersistence();
    el.generate.addEventListener("click", onGenerateClicked);
  }

  function renderSummary() {
    var item = currentItem();
    el.subject.textContent = item.subject || "(kein Betreff)";
    el.from.textContent = formatAddress(item.from);
    el.date.textContent = formatDate(item.dateTimeCreated);
  }

  function renderAttachmentList() {
    var item = currentItem();
    currentAttachments = nonInlineAttachments(item);

    el.attachmentList.innerHTML = "";

    if (currentAttachments.length === 0) {
      el.noAttachments.hidden = false;
      return;
    }
    el.noAttachments.hidden = true;

    currentAttachments.forEach(function (attachment, index) {
      var li = document.createElement("li");
      li.className = "mp-attachment-item";

      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = true;
      checkbox.id = "mp-att-" + index;
      checkbox.dataset.attachmentIndex = String(index);

      var label = document.createElement("label");
      label.className = "mp-attachment-meta";
      label.setAttribute("for", checkbox.id);

      var nameSpan = document.createElement("span");
      nameSpan.className = "mp-attachment-name";
      nameSpan.textContent = attachment.name || "(unbenannt)";

      var detailSpan = document.createElement("span");
      detailSpan.className = "mp-attachment-detail";
      detailSpan.textContent =
        (attachment.contentType || "unbekannter Typ") + " · " + Postfach2PdfCore.formatBytes(attachment.size);

      label.appendChild(nameSpan);
      label.appendChild(detailSpan);
      li.appendChild(checkbox);
      li.appendChild(label);
      el.attachmentList.appendChild(li);
    });
  }

  function getSelectedAttachments() {
    return currentAttachments.filter(function (attachment, index) {
      var checkbox = document.getElementById("mp-att-" + index);
      return checkbox ? checkbox.checked : true;
    });
  }

  function getDeselectedAttachments() {
    return currentAttachments.filter(function (attachment, index) {
      var checkbox = document.getElementById("mp-att-" + index);
      return checkbox ? !checkbox.checked : false;
    });
  }

  function loadPersistedSettings() {
    var settings = Office.context.roamingSettings;
    if (!settings) {
      el.pageNumbers.checked = true;
      return;
    }
    el.loadExternalImages.checked = settings.get(SETTINGS_KEYS.loadExternalImages) === true;
    el.includeHeaders.checked = settings.get(SETTINGS_KEYS.includeHeaders) === true;
    var storedPageNumbers = settings.get(SETTINGS_KEYS.pageNumbers);
    el.pageNumbers.checked = storedPageNumbers === undefined ? true : storedPageNumbers === true;
  }

  function persistSetting(key, value) {
    var settings = Office.context.roamingSettings;
    if (!settings) {
      return;
    }
    settings.set(key, value);
    settings.saveAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("Postfach2PDF: Einstellung konnte nicht gespeichert werden", result.error);
      }
    });
  }

  function wireSettingsPersistence() {
    el.loadExternalImages.addEventListener("change", function () {
      persistSetting(SETTINGS_KEYS.loadExternalImages, el.loadExternalImages.checked);
    });
    el.includeHeaders.addEventListener("change", function () {
      persistSetting(SETTINGS_KEYS.includeHeaders, el.includeHeaders.checked);
    });
    el.pageNumbers.addEventListener("change", function () {
      persistSetting(SETTINGS_KEYS.pageNumbers, el.pageNumbers.checked);
    });
  }

  function buildFileName(item) {
    var datePart = item.dateTimeCreated
      ? item.dateTimeCreated.toISOString().slice(0, 10)
      : "unbekannt-datum";
    var senderPart =
      Postfach2PdfCore.sanitizeFileNamePart(
        (item.from && (item.from.displayName || item.from.emailAddress)) || "unbekannt"
      ) || "unbekannt";
    var subjectPart = Postfach2PdfCore.sanitizeFileNamePart(item.subject || "kein Betreff") || "kein Betreff";
    var stem = Postfach2PdfCore.sanitizeFileNameStem(datePart + "_" + senderPart + "_" + subjectPart);
    return stem + ".pdf";
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function stripExternalImages(html) {
    // Ersetzt src-Attribute von <img>, die nicht bereits als data:-URI
    // eingebettet sind, durch einen Platzhalter, damit standardmaessig keine
    // Remote-Ressourcen (z. B. Tracking-Pixel) nachgeladen werden.
    return html.replace(/<img\b([^>]*?)\ssrc=(["'])(.*?)\2([^>]*)>/gi, function (
      match,
      before,
      quote,
      src,
      after
    ) {
      if (/^data:/i.test(src)) {
        return match;
      }
      return (
        '<img' +
        before +
        ' src="" data-mailpdf-blocked-src="' +
        escapeHtml(src) +
        '" alt="[externes Bild blockiert]"' +
        after +
        ">"
      );
    });
  }

  function getBodyHtml(item) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function buildDocumentHtml(item, bodyHtml, options, selectedAttachments) {
    var headerRows = [
      ["Betreff", escapeHtml(item.subject || "(kein Betreff)")],
      ["Von", escapeHtml(formatAddress(item.from))],
      ["An", escapeHtml(formatAddressList(item.to))],
    ];
    if (item.cc && item.cc.length > 0) {
      headerRows.push(["CC", escapeHtml(formatAddressList(item.cc))]);
    }
    headerRows.push(["Datum", escapeHtml(formatDate(item.dateTimeCreated))]);

    if (options.includeHeaders) {
      if (item.internetMessageId) {
        headerRows.push(["Nachrichten-ID", escapeHtml(item.internetMessageId)]);
      }
      if (item.conversationId) {
        headerRows.push(["Konversations-ID", escapeHtml(item.conversationId)]);
      }
    }

    if (selectedAttachments.length > 0) {
      headerRows.push([
        "Anhaenge",
        escapeHtml(selectedAttachments.map(function (a) { return a.name; }).join(", ")),
      ]);
    }

    var headerHtml = headerRows
      .map(function (row) {
        return (
          '<tr><td class="mp-pdf-label">' +
          row[0] +
          '</td><td class="mp-pdf-value">' +
          row[1] +
          "</td></tr>"
        );
      })
      .join("");

    var processedBody = options.loadExternalImages ? bodyHtml : stripExternalImages(bodyHtml);

    return (
      '<div class="mp-pdf-document">' +
      '<h1 class="mp-pdf-subject">' +
      escapeHtml(item.subject || "(kein Betreff)") +
      "</h1>" +
      '<table class="mp-pdf-header">' +
      headerHtml +
      "</table>" +
      '<hr class="mp-pdf-divider" />' +
      '<div class="mp-pdf-body">' +
      processedBody +
      "</div>" +
      "</div>"
    );
  }

  var PDF_DOCUMENT_STYLE =
    "<style>" +
    ".mp-pdf-document { font-family: Calibri, Arial, sans-serif; font-size: 12px; color: #000; }" +
    ".mp-pdf-subject { font-size: 18px; margin: 0 0 10px 0; }" +
    ".mp-pdf-header td { padding: 2px 8px 2px 0; vertical-align: top; }" +
    ".mp-pdf-label { color: #555; white-space: nowrap; font-weight: 600; }" +
    ".mp-pdf-divider { border: none; border-top: 1px solid #999; margin: 10px 0; }" +
    ".mp-pdf-body img { max-width: 100%; height: auto; }" +
    ".mp-pdf-body table { max-width: 100%; }" +
    "</style>";

  // Muss mit der Breite von .mp-render-root in taskpane.css uebereinstimmen.
  var RENDER_ROOT_WIDTH = 800;

  async function renderEmailHtmlToPdfBytes(item, bodyHtml, options, attachmentsForHeader) {
    var documentHtml = buildDocumentHtml(item, bodyHtml, options, attachmentsForHeader);
    el.renderRoot.innerHTML = PDF_DOCUMENT_STYLE + documentHtml;
    try {
      return await html2pdf()
        .set({
          margin: [15, 12, 15, 12],
          // windowWidth/width/x/y explizit setzen: ohne diese Angaben
          // berechnet diese html2canvas-Version einen falschen
          // horizontalen Ausschnitt (fester Versatz von ca. 259px,
          // unabhaengig vom Inhalt) und schneidet den linken Rand +
          // Teile des Textes ab. Lokal mit Playwright/Chromium gegen
          // die vendorte Bibliothek reproduziert und verifiziert.
          html2canvas: {
            scale: 2,
            useCORS: false,
            windowWidth: RENDER_ROOT_WIDTH,
            width: RENDER_ROOT_WIDTH,
            x: 0,
            y: 0,
          },
          jsPDF: { unit: "pt", format: "a4", orientation: "portrait" },
          pagebreak: { mode: ["css", "legacy"] },
        })
        .from(el.renderRoot)
        .toPdf()
        .outputPdf("arraybuffer");
    } finally {
      el.renderRoot.innerHTML = "";
    }
  }

  // Rendert eine einzelne E-Mail (Body + alle Anhaenge) als Seiten in das
  // uebergebene PDFDocument. Wird sowohl vom Einzel- als auch vom
  // Batch-Modus genutzt.
  async function renderEmailPagesInto(item, options, attachments, targetPdf, results, onProgress) {
    var bodyHtml = await getBodyHtml(item);
    var bodyPdfArrayBuffer = await renderEmailHtmlToPdfBytes(item, bodyHtml, options, attachments);

    var bodyPdf = await PDFLib.PDFDocument.load(bodyPdfArrayBuffer);
    var bodyPages = await targetPdf.copyPages(bodyPdf, bodyPdf.getPageIndices());
    bodyPages.forEach(function (p) {
      targetPdf.addPage(p);
    });

    for (var i = 0; i < attachments.length; i++) {
      var attachment = attachments[i];
      if (onProgress) {
        onProgress(attachment, i, attachments.length);
      }
      var outcome = await Postfach2PdfCore.embedAttachmentIntoPdf(item, targetPdf, attachment);
      results.push({ name: attachment.name, status: outcome.status, reason: outcome.reason });
    }
  }

  async function buildFinalPdf(item, options) {
    var selectedAttachments = getSelectedAttachments();
    var deselectedAttachments = getDeselectedAttachments();
    var results = [];

    var mergedPdf = await PDFLib.PDFDocument.create();
    mergedPdf.setProducer("Postfach2PDF (schantastic)");
    mergedPdf.setCreator("Postfach2PDF");
    mergedPdf.setTitle(item.subject || "E-Mail");
    mergedPdf.setCreationDate(new Date());

    await renderEmailPagesInto(item, options, selectedAttachments, mergedPdf, results, function (attachment, i, total) {
      setStatus("Verarbeite Anhang " + (i + 1) + "/" + total + ": " + attachment.name + " ...");
    });

    deselectedAttachments.forEach(function (attachment) {
      results.push({ name: attachment.name, status: "skipped" });
    });

    if (options.pageNumbers) {
      setStatus("Fuege Seitenzahlen hinzu ...");
      await Postfach2PdfCore.addPageNumbers(mergedPdf);
    }

    setStatus("Erzeuge PDF-Bytes ...");
    var finalBytes = await mergedPdf.save();
    var fileName = buildFileName(item);

    return { bytes: finalBytes, fileName: fileName, results: results };
  }

  function currentOptions() {
    return {
      loadExternalImages: el.loadExternalImages.checked,
      includeHeaders: el.includeHeaders.checked,
      pageNumbers: el.pageNumbers.checked,
    };
  }

  async function onGenerateClicked() {
    el.generate.disabled = true;
    el.result.hidden = true;
    setStatus("Erzeuge PDF ...");

    var item = currentItem();

    try {
      var built = await buildFinalPdf(item, currentOptions());
      setStatus("Speichere PDF ...");
      Postfach2PdfCore.downloadBytes(built.bytes, built.fileName);
      renderResults(built.results);
      setStatus("PDF wurde erzeugt und zum Download angeboten.", "success");
    } catch (error) {
      console.error("Postfach2PDF: PDF-Erzeugung fehlgeschlagen", error);
      setStatus(
        "Fehler bei der PDF-Erzeugung: " + (error && error.message ? error.message : error),
        "error"
      );
    } finally {
      el.generate.disabled = false;
    }
  }

  // =======================================================================
  // Batch-Modus (Mehrfachauswahl mehrerer E-Mails)
  // =======================================================================

  function getSelectedItemsAsync() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.getSelectedItemsAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function loadItemByIdAsync(itemId) {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.loadItemByIdAsync(itemId, {}, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  async function initBatchMode() {
    debugLog("initBatchMode gestartet (kein einzelnes Element aktiv, Mailbox 1.13 unterstuetzt)");
    el.batch.hidden = false;
    el.batchStart.addEventListener("click", onBatchStartClicked);

    // Laut Microsoft-Doku zum Mehrfachauswahl-Feature aktualisiert sich
    // die Auswahl waehrend die Taskpane offen bleibt nur ueber dieses
    // Event, nicht automatisch.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.SelectedItemsChanged,
      function () {
        debugLog("SelectedItemsChanged-Event ausgeloest");
        refreshBatchSelection();
      },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          debugLog("SelectedItemsChanged-Handler konnte nicht registriert werden: " + JSON.stringify(result.error));
        } else {
          debugLog("SelectedItemsChanged-Handler erfolgreich registriert");
        }
      }
    );

    await refreshBatchSelectionWithRetry();
  }

  async function refreshBatchSelection() {
    try {
      batchSelection = await getSelectedItemsAsync();
    } catch (error) {
      debugLog("Mehrfachauswahl konnte nicht gelesen werden: " + JSON.stringify(error));
      batchSelection = [];
    }

    renderBatchList();
  }

  // Direkt nach dem Laden der Taskpane kann getSelectedItemsAsync kurzzeitig
  // leer zurueckkommen, weil Outlook die Auswahl noch nicht an die Taskpane
  // durchgereicht hat (das offizielle Microsoft-Beispiel verlaesst sich
  // deshalb ausschliesslich auf das SelectedItemsChanged-Event statt auf
  // einen sofortigen Aufruf). Hier stattdessen ein paar Mal mit kurzer
  // Verzoegerung erneut versuchen, bevor "keine Auswahl" angezeigt wird.
  async function refreshBatchSelectionWithRetry() {
    var retryDelaysMs = [0, 300, 600, 1200];

    for (var i = 0; i < retryDelaysMs.length; i++) {
      if (retryDelaysMs[i] > 0) {
        await new Promise(function (resolve) {
          setTimeout(resolve, retryDelaysMs[i]);
        });
      }

      try {
        batchSelection = await getSelectedItemsAsync();
        debugLog(
          "getSelectedItemsAsync (Versuch " + (i + 1) + "/" + retryDelaysMs.length + ") ergab " +
            batchSelection.length + " Element(e): " +
            JSON.stringify(
              batchSelection.map(function (m) {
                return { itemType: m.itemType, itemMode: m.itemMode, subject: m.subject };
              })
            )
        );
      } catch (error) {
        debugLog("Mehrfachauswahl konnte nicht gelesen werden: " + JSON.stringify(error));
        batchSelection = [];
      }

      var hasMessages = batchSelection.some(function (m) {
        return m.itemType === Office.MailboxEnums.ItemType.Message;
      });
      if (hasMessages) {
        break;
      }
    }

    renderBatchList();
  }

  function renderBatchList() {
    el.batchList.innerHTML = "";

    var messages = batchSelection.filter(function (m) {
      return m.itemType === Office.MailboxEnums.ItemType.Message;
    });

    if (messages.length === 0) {
      el.batchEmpty.hidden = false;
      el.batchStart.disabled = true;
      return;
    }
    el.batchEmpty.hidden = true;
    el.batchStart.disabled = false;

    messages.forEach(function (message, index) {
      var li = document.createElement("li");
      li.className = "mp-attachment-item";

      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = true;
      checkbox.id = "mp-batch-item-" + index;
      checkbox.dataset.batchIndex = String(index);

      var label = document.createElement("label");
      label.className = "mp-attachment-meta";
      label.setAttribute("for", checkbox.id);

      var nameSpan = document.createElement("span");
      nameSpan.className = "mp-attachment-name";
      nameSpan.textContent = message.subject || "(kein Betreff)";

      var detailSpan = document.createElement("span");
      detailSpan.className = "mp-attachment-detail";
      detailSpan.textContent = message.hasAttachments ? "hat Anhaenge" : "keine Anhaenge";

      label.appendChild(nameSpan);
      label.appendChild(detailSpan);
      li.appendChild(checkbox);
      li.appendChild(label);
      el.batchList.appendChild(li);
    });

    batchSelection = messages;
  }

  function getSelectedBatchMessages() {
    return batchSelection.filter(function (message, index) {
      var checkbox = document.getElementById("mp-batch-item-" + index);
      return checkbox ? checkbox.checked : true;
    });
  }

  function batchOptions() {
    return {
      loadExternalImages: false,
      includeHeaders: el.batchIncludeHeaders.checked,
      pageNumbers: el.batchPageNumbers.checked,
    };
  }

  function buildBatchFileName(messages) {
    var stem = Postfach2PdfCore.sanitizeFileNameStem(
      new Date().toISOString().slice(0, 10) + "_Postfach2PDF-Batch_" + messages.length + "-Mails"
    );
    return stem + ".pdf";
  }

  async function convertBatchMerged(messages, options) {
    var mergedPdf = await PDFLib.PDFDocument.create();
    mergedPdf.setProducer("Postfach2PDF (schantastic)");
    mergedPdf.setCreator("Postfach2PDF");
    mergedPdf.setCreationDate(new Date());
    var overallResults = [];

    for (var i = 0; i < messages.length; i++) {
      setStatus("Lade E-Mail " + (i + 1) + "/" + messages.length + ": " + (messages[i].subject || "") + " ...");
      var item = await loadItemByIdAsync(messages[i].itemId);
      var attachments = nonInlineAttachments(item);
      var itemResults = [];
      await renderEmailPagesInto(item, options, attachments, mergedPdf, itemResults, function (attachment, j, total) {
        setStatus(
          "E-Mail " + (i + 1) + "/" + messages.length + " – Anhang " + (j + 1) + "/" + total + ": " + attachment.name + " ..."
        );
      });
      overallResults.push({
        name: messages[i].subject || "(kein Betreff)",
        status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
        reason: "in gemeinsame PDF eingebettet (" + itemResults.length + " Anhang-Ergebnisse)",
      });
    }

    if (options.pageNumbers) {
      setStatus("Fuege Seitenzahlen hinzu ...");
      await Postfach2PdfCore.addPageNumbers(mergedPdf);
    }

    var bytes = await mergedPdf.save();
    var fileName = buildBatchFileName(messages);
    Postfach2PdfCore.downloadBytes(bytes, fileName);
    return overallResults;
  }

  async function convertBatchSeparate(messages, options) {
    var overallResults = [];

    for (var i = 0; i < messages.length; i++) {
      setStatus("Lade E-Mail " + (i + 1) + "/" + messages.length + ": " + (messages[i].subject || "") + " ...");
      var item = await loadItemByIdAsync(messages[i].itemId);
      var attachments = nonInlineAttachments(item);
      var itemResults = [];

      var singleDoc = await PDFLib.PDFDocument.create();
      singleDoc.setProducer("Postfach2PDF (schantastic)");
      singleDoc.setCreator("Postfach2PDF");
      singleDoc.setTitle(item.subject || "E-Mail");
      singleDoc.setCreationDate(new Date());

      await renderEmailPagesInto(item, options, attachments, singleDoc, itemResults, function (attachment, j, total) {
        setStatus(
          "E-Mail " + (i + 1) + "/" + messages.length + " – Anhang " + (j + 1) + "/" + total + ": " + attachment.name + " ..."
        );
      });

      if (options.pageNumbers) {
        await Postfach2PdfCore.addPageNumbers(singleDoc);
      }

      var bytes = await singleDoc.save();
      var fileName = buildFileName(item);
      Postfach2PdfCore.downloadBytes(bytes, fileName);

      overallResults.push({
        name: messages[i].subject || "(kein Betreff)",
        status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
        reason: "als eigene PDF (" + fileName + ") gespeichert",
      });

      // Kurze Pause, damit Browser mehrere aufeinanderfolgende Downloads
      // zuverlaessiger zulassen.
      await new Promise(function (resolve) {
        setTimeout(resolve, 400);
      });
    }

    return overallResults;
  }

  async function onBatchStartClicked() {
    el.batchStart.disabled = true;
    el.result.hidden = true;
    setStatus("Starte Batch-Verarbeitung ...");

    var messages = getSelectedBatchMessages();
    if (messages.length === 0) {
      setStatus("Keine E-Mails ausgewaehlt.", "error");
      el.batchStart.disabled = false;
      return;
    }

    var options = batchOptions();

    try {
      var results;
      if (el.batchMerge.checked) {
        results = await convertBatchMerged(messages, options);
      } else {
        results = await convertBatchSeparate(messages, options);
      }
      renderResults(results);
      var failedCount = results.filter(function (r) {
        return r.status === "failed";
      }).length;
      if (failedCount === 0) {
        setStatus(messages.length + " E-Mail(s) erfolgreich verarbeitet.", "success");
      } else {
        setStatus(failedCount + " von " + messages.length + " E-Mails hatten Probleme, siehe Ergebnis.", "error");
      }
    } catch (error) {
      console.error("Postfach2PDF: Batch-Verarbeitung fehlgeschlagen", error);
      setStatus(
        "Fehler bei der Batch-Verarbeitung: " + (error && error.message ? error.message : error),
        "error"
      );
    } finally {
      el.batchStart.disabled = false;
    }
  }
})();
