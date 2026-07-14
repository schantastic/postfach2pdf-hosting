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

    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      el.unsupported.hidden = false;
      el.unsupported.textContent =
        "Dieses Outlook unterstuetzt die fuer Postfach2PDF benoetigte Mailbox-API (1.8, fuer Anhangszugriff) nicht.";
      return;
    }

    el.shared.hidden = false;

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

  // Eingebettete (Inline-)Bilder - z. B. Signaturlogos, im Body per
  // "cid:..." oder (in aelteren Outlook-Versionen) per interner,
  // authentifizierter URL referenziert - sind in Wahrheit ganz normale
  // Anhaenge mit isInline=true, nur nicht in der Anhangsliste sichtbar.
  // Ein direkter fetch() auf die Body-URL schlaegt bei diesen IMMER fehl
  // (0 von 13 in einem realen Testfall), weil diese URLs eine
  // Outlook-Session brauchen, auf die unsere Taskpane keinen Zugriff hat.
  // Der korrekte Weg ist Office.js' eigene, bereits authentifizierte
  // getAttachmentContentAsync-API (wie fuer normale Anhaenge) statt die
  // URL selbst abzurufen. Zuordnung ueber contentId (neueres "cid:..."-
  // Format, Mailbox 1.16) oder ueber die Anhang-ID als Teilstring der src
  // (aelteres Format, in dem die Anhang-ID direkt in der URL steht).
  async function embedInlineAttachmentImages(item, container) {
    var inlineAttachments = (item.attachments || []).filter(function (a) {
      return a.isInline;
    });
    var stats = { total: inlineAttachments.length, embedded: 0, failed: 0 };
    if (inlineAttachments.length === 0) {
      return stats;
    }

    var imgs = Array.prototype.slice.call(container.querySelectorAll("img"));

    await Promise.all(
      inlineAttachments.map(async function (attachment) {
        try {
          var content = await Postfach2PdfCore.getAttachmentContent(item, attachment.id);
          if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
            throw new Error("unerwartetes Format " + content.format);
          }
          var dataUrl = "data:" + (attachment.contentType || "image/png") + ";base64," + content.content;
          var matched = false;
          imgs.forEach(function (img) {
            var src = img.getAttribute("src") || "";
            var matchesCid = attachment.contentId && src === "cid:" + attachment.contentId;
            var matchesId = src.indexOf(attachment.id) !== -1;
            if (matchesCid || matchesId) {
              img.src = dataUrl;
              matched = true;
            }
          });
          if (matched) {
            stats.embedded++;
          } else {
            stats.failed++;
          }
        } catch (error) {
          stats.failed++;
          console.warn("Postfach2PDF: Inline-Bild konnte nicht eingebettet werden", attachment.name, error);
        }
      })
    );

    return stats;
  }

  // Laedt externe (http/https) Bilder selbst herunter und ersetzt src durch
  // eine data:-URI, BEVOR html2canvas ueberhaupt anfaengt. Grund: mit
  // Playwright/Chromium reproduziert - html2canvas klont das DOM fuer die
  // eigentliche Erfassung intern nochmal, und dieser Klon wartet nicht
  // zuverlaessig auf noch ladende Netzwerkbilder. Ergebnis war ein Bild mit
  // ~90px statt der echten Hoehe und kompletter Verlust allen Inhalts
  // danach - selbst wenn das Original-<img> im sichtbaren DOM laengst fertig
  // geladen war. Data-URIs sind synchron/sofort verfuegbar und umgehen das
  // Problem komplett. Schlaegt der Download fehl (z. B. CORS - erwartbar
  // fuer echte Outlook-interne Bild-URLs, siehe embedInlineAttachmentImages
  // oben, die das eigentlich schon vorher abfangen sollte), bleibt das
  // Bild als normaler Live-Link stehen (gleiches Verhalten wie vorher).
  async function inlineExternalImages(container) {
    var imgs = Array.prototype.slice.call(container.querySelectorAll("img"));
    var stats = { total: 0, inlined: 0, failed: 0, firstError: null };

    await Promise.all(
      imgs.map(function (img) {
        var src = img.getAttribute("src") || "";
        if (!/^https?:/i.test(src)) {
          return Promise.resolve();
        }
        stats.total++;
        return fetch(src)
          .then(function (response) {
            if (!response.ok) {
              throw new Error("HTTP " + response.status);
            }
            return response.blob();
          })
          .then(function (blob) {
            return new Promise(function (resolve, reject) {
              var reader = new FileReader();
              reader.onload = function () {
                resolve(reader.result);
              };
              reader.onerror = function () {
                reject(reader.error);
              };
              reader.readAsDataURL(blob);
            });
          })
          .then(function (dataUrl) {
            img.src = dataUrl;
            stats.inlined++;
          })
          .catch(function (error) {
            stats.failed++;
            // Nur den Hostnamen merken (nicht die volle URL mit evtl.
            // Tracking-Tokens) + die Fehlermeldung, damit man ohne
            // DevTools sieht, WARUM der Download scheitert (CORS? HTTP-
            // Status? Netzwerkfehler? Browser-eigener Tracker-/Werbeblocker
            // wie Brave Shields?), ohne sensible URL-Teile preiszugeben.
            if (!stats.firstError) {
              var hostname = src;
              try {
                hostname = new URL(src).hostname;
              } catch (e) {
                // src war keine gueltige absolute URL - Rohwert behalten
              }
              stats.firstError = hostname + ": " + (error && error.message ? error.message : String(error));
            }
            // Bild NICHT als lebendigen Netzwerk-Link stehen lassen: genau
            // das fuehrt dazu, dass html2canvas beim internen DOM-Klonen
            // haengen bleibt und ALLEN Inhalt danach verliert (reproduziert
            // und in STATUS.md dokumentiert). Lieber ein klar erkennbarer
            // Platzhalter als kompletter Content-Verlust ab dieser Stelle.
            img.removeAttribute("src");
            img.alt = "[Bild konnte nicht geladen werden]";
            console.warn(
              "Postfach2PDF: externes Bild konnte nicht eingebettet werden, wird als Platzhalter entfernt",
              src,
              error
            );
          });
      })
    );

    return stats;
  }

  // Wartet, bis alle <img>-Elemente im Container fertig geladen (oder
  // endgueltig fehlgeschlagen) sind. Faengt vor allem Bilder ab, die trotz
  // inlineExternalImages() noch als Live-Link uebrig sind (z. B. wegen
  // CORS). Pro Bild maximal 8s warten, damit ein einzelnes haengendes Bild
  // die PDF-Erzeugung nicht blockiert.
  function waitForImages(container) {
    var imgs = container.querySelectorAll("img");
    var promises = [];
    imgs.forEach(function (img) {
      if (img.complete) {
        return;
      }
      promises.push(
        new Promise(function (resolve) {
          var done = function () {
            img.removeEventListener("load", done);
            img.removeEventListener("error", done);
            resolve();
          };
          img.addEventListener("load", done);
          img.addEventListener("error", done);
          setTimeout(done, 8000);
        })
      );
    });
    return Promise.all(promises);
  }

  async function renderEmailHtmlToPdfBytes(item, bodyHtml, options, attachmentsForHeader) {
    var documentHtml = buildDocumentHtml(item, bodyHtml, options, attachmentsForHeader);
    el.renderRoot.innerHTML = PDF_DOCUMENT_STYLE + documentHtml;
    var inlineImageStats = await embedInlineAttachmentImages(item, el.renderRoot);
    var imageStats = await inlineExternalImages(el.renderRoot);
    await waitForImages(el.renderRoot);
    // Kurze Pause, bevor html2canvas zu erfassen beginnt: Nutzer meldete
    // fehlenden Inhalt ganz am Ende grosser Mails, der sich trotz
    // identischem HTML/Bilder-Zustand lokal nicht reproduzieren liess -
    // vermutlich ein Timing-/Ressourcenproblem, das nur in der echten,
    // sehr schweren Outlook-Webseite auftritt (viel eigenes JavaScript
    // parallel zu unserer Taskpane). Alle bisher in diesem Projekt
    // gefundenen Rendering-Bugs waren Timing-bedingt - diese Pause gibt
    // Layout/GC zusaetzliche Zeit, sich vor der aufwendigen Erfassung
    // zu setzen.
    await new Promise(function (resolve) {
      setTimeout(resolve, 300);
    });
    // Fusszeile/Seitenzahl (addPageNumbers) braucht ~30pt Platz am unteren
    // Rand jeder Seite. Bei 15pt Rand ueberlappt der letzte Textabschnitt
    // einer vollen Seite die Fusszeile (mit echten Chromium-Renderings
    // reproduziert und verifiziert) - deshalb mehr Rand, wenn Seitenzahlen
    // aktiv sind.
    var bottomMargin = options.pageNumbers ? 45 : 15;
    // Temporaer zur Fehlersuche (Content-Verlust bei grossen HTML-Mails):
    // Hoehe des fertig aufgebauten Inhalts VOR dem html2canvas-Erfassen
    // messen, um zu sehen ob das DOM selbst schon zu kurz ist (Problem im
    // Mail-HTML/CSS) oder ob html2canvas/html2pdf beim Erfassen etwas
    // verliert (Problem in der Rendering-Pipeline).
    var renderRootScrollHeight = el.renderRoot.scrollHeight;
    try {
      var arrayBuffer = await html2pdf()
        .set({
          margin: [15, 12, bottomMargin, 12],
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
      return {
        arrayBuffer: arrayBuffer,
        renderRootScrollHeight: renderRootScrollHeight,
        imageStats: imageStats,
        inlineImageStats: inlineImageStats,
      };
    } finally {
      el.renderRoot.innerHTML = "";
    }
  }

  // Rendert eine einzelne E-Mail (Body + alle Anhaenge) als Seiten in das
  // uebergebene PDFDocument. Wird sowohl vom Einzel- als auch vom
  // Batch-Modus genutzt. skipPageIndices sammelt die 0-basierten Indizes
  // seitengetreu kopierter PDF-Anhangsseiten (kein reservierter Rand) -
  // addPageNumbers() darf darauf nichts zeichnen, siehe dort.
  // Temporaer zur Fehlersuche: haelt das zuletzt von Office.js erhaltene
  // Roh-HTML fest, damit onGenerateClicked es zusaetzlich als .txt zum
  // Download anbieten kann (der rohe .eml-Quelltext unterscheidet sich
  // von dem, was Office.js tatsaechlich liefert - Outlook schreibt u. a.
  // Bild-URLs um - deshalb reicht ein lokaler .eml-Test nicht).
  var lastBodyHtmlForDebug = null;

  async function renderEmailPagesInto(item, options, attachments, targetPdf, results, skipPageIndices, onProgress) {
    var bodyHtml = await getBodyHtml(item);
    lastBodyHtmlForDebug = bodyHtml;
    var rendered = await renderEmailHtmlToPdfBytes(item, bodyHtml, options, attachments);

    var bodyPdf = await PDFLib.PDFDocument.load(rendered.arrayBuffer);
    // Temporaer zur Fehlersuche (Content-Verlust bei grossen HTML-Mails).
    // Erscheint in der Ergebnisliste, damit man ohne DevTools sieht: wie
    // viele Zeichen kamen von Office.js an, wie hoch (px) war der fertig
    // aufgebaute Inhalt VOR html2canvas, und wie viele Seiten kamen dabei
    // raus. Grosse Hoehe + wenige Seiten = Rendering verliert Inhalt.
    // Kleine Hoehe trotz vieler Zeichen = Problem liegt im Mail-HTML/CSS
    // selbst (z. B. overflow:hidden in zusammengeklapptem Zitat-Verlauf).
    results.push({
      name: "Diagnose: Body-Laenge",
      status: "embedded",
      detail:
        bodyHtml.length + " Zeichen HTML, Render-Hoehe " + rendered.renderRootScrollHeight +
        "px, " + bodyPdf.getPageCount() + " Body-Seite(n) gerendert. Inline-Anhangsbilder: " +
        rendered.inlineImageStats.embedded + " eingebettet / " + rendered.inlineImageStats.failed +
        " fehlgeschlagen / " + rendered.inlineImageStats.total + " gesamt. Externe Bild-Downloads: " +
        rendered.imageStats.inlined + " eingebettet / " + rendered.imageStats.failed +
        " fehlgeschlagen / " + rendered.imageStats.total + " gesamt" +
        (rendered.imageStats.firstError ? ". Erster Fehler: " + rendered.imageStats.firstError : ""),
    });
    // Temporaer zur Fehlersuche: zeigt die letzten ~300 Zeichen des von
    // Office.js erhaltenen Roh-HTML (als Text, Tags entfernt), damit man
    // sieht ob der erwartete Fusszeilentext (z. B. "Unsubscribe") darin
    // ueberhaupt vorkommt und wo - hilft zu unterscheiden, ob Office.js
    // schon unvollstaendig liefert oder ob es erst beim Rendern verloren
    // geht.
    results.push({
      name: "Diagnose: HTML-Ende (roh)",
      status: "embedded",
      detail: "..." + bodyHtml.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").slice(-300),
    });
    var bodyPages = await targetPdf.copyPages(bodyPdf, bodyPdf.getPageIndices());
    bodyPages.forEach(function (p) {
      targetPdf.addPage(p);
    });

    for (var i = 0; i < attachments.length; i++) {
      var attachment = attachments[i];
      if (onProgress) {
        onProgress(attachment, i, attachments.length);
      }
      var pageCountBefore = targetPdf.getPageCount();
      var outcome = await Postfach2PdfCore.embedAttachmentIntoPdf(item, targetPdf, attachment);
      if (outcome.rawCopy) {
        var pageCountAfter = targetPdf.getPageCount();
        for (var p = pageCountBefore; p < pageCountAfter; p++) {
          skipPageIndices.push(p);
        }
      }
      results.push({ name: attachment.name, status: outcome.status, reason: outcome.reason });
    }
  }

  async function buildFinalPdf(item, options) {
    var selectedAttachments = getSelectedAttachments();
    var deselectedAttachments = getDeselectedAttachments();
    var results = [];
    var skipPageIndices = [];

    var mergedPdf = await PDFLib.PDFDocument.create();
    mergedPdf.setProducer("Postfach2PDF (schantastic)");
    mergedPdf.setCreator("Postfach2PDF");
    mergedPdf.setTitle(item.subject || "E-Mail");
    mergedPdf.setCreationDate(new Date());

    await renderEmailPagesInto(item, options, selectedAttachments, mergedPdf, results, skipPageIndices, function (attachment, i, total) {
      setStatus("Verarbeite Anhang " + (i + 1) + "/" + total + ": " + attachment.name + " ...");
    });

    deselectedAttachments.forEach(function (attachment) {
      results.push({ name: attachment.name, status: "skipped" });
    });

    if (options.pageNumbers) {
      setStatus("Fuege Seitenzahlen hinzu ...");
      await Postfach2PdfCore.addPageNumbers(mergedPdf, skipPageIndices);
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
      // Temporaer zur Fehlersuche: rohes Body-HTML zusaetzlich als .txt
      // anbieten (siehe lastBodyHtmlForDebug oben).
      if (lastBodyHtmlForDebug) {
        var debugBytes = new TextEncoder().encode(lastBodyHtmlForDebug);
        Postfach2PdfCore.downloadBytes(debugBytes, "postfach2pdf-debug-body.txt");
      }
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

  // Laut Microsoft-Doku darf immer nur eine per loadItemByIdAsync geladene
  // Nachricht gleichzeitig geladen sein - ohne unloadAsync vor dem naechsten
  // loadItemByIdAsync schlaegt der Batch-Modus ab der zweiten Mail fehl
  // ("Das Element ist nicht vorhanden oder wurde nicht erstellt.").
  function unloadItemAsync(item) {
    return new Promise(function (resolve) {
      item.unloadAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Postfach2PDF: unloadAsync fehlgeschlagen", result.error);
        }
        resolve();
      });
    });
  }

  async function initBatchMode() {
    el.batch.hidden = false;
    el.batchStart.addEventListener("click", onBatchStartClicked);

    // Laut Microsoft-Doku zum Mehrfachauswahl-Feature aktualisiert sich
    // die Auswahl waehrend die Taskpane offen bleibt nur ueber dieses
    // Event, nicht automatisch.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.SelectedItemsChanged,
      refreshBatchSelection,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Postfach2PDF: SelectedItemsChanged-Handler konnte nicht registriert werden", result.error);
        }
      }
    );

    await refreshBatchSelectionWithRetry();
  }

  async function refreshBatchSelection() {
    try {
      batchSelection = await getSelectedItemsAsync();
    } catch (error) {
      console.error("Postfach2PDF: Mehrfachauswahl konnte nicht gelesen werden", error);
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
      } catch (error) {
        console.error("Postfach2PDF: Mehrfachauswahl konnte nicht gelesen werden", error);
        batchSelection = [];
      }

      var hasMessages = batchSelection.some(isMessageItem);
      if (hasMessages) {
        break;
      }
    }

    renderBatchList();
  }

  // Diagnose zeigte: getSelectedItemsAsync liefert Elemente mit
  // itemType "Message" (String), der direkte Vergleich gegen
  // Office.MailboxEnums.ItemType.Message matchte trotzdem nicht - deshalb
  // hier zusaetzlich gegen den literalen String verglichen statt sich
  // allein auf den Enum-Wert zu verlassen.
  function isMessageItem(m) {
    return m.itemType === Office.MailboxEnums.ItemType.Message || m.itemType === "Message";
  }

  function renderBatchList() {
    el.batchList.innerHTML = "";

    var messages = batchSelection.filter(isMessageItem);

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
    var skipPageIndices = [];

    for (var i = 0; i < messages.length; i++) {
      setStatus("Lade E-Mail " + (i + 1) + "/" + messages.length + ": " + (messages[i].subject || "") + " ...");
      var item = await loadItemByIdAsync(messages[i].itemId);
      try {
        var attachments = nonInlineAttachments(item);
        var itemResults = [];
        await renderEmailPagesInto(item, options, attachments, mergedPdf, itemResults, skipPageIndices, function (attachment, j, total) {
          setStatus(
            "E-Mail " + (i + 1) + "/" + messages.length + " – Anhang " + (j + 1) + "/" + total + ": " + attachment.name + " ..."
          );
        });
        overallResults.push({
          name: messages[i].subject || "(kein Betreff)",
          status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
          reason: "in gemeinsame PDF eingebettet (" + itemResults.length + " Anhang-Ergebnisse)",
        });
      } finally {
        await unloadItemAsync(item);
      }
    }

    if (options.pageNumbers) {
      setStatus("Fuege Seitenzahlen hinzu ...");
      await Postfach2PdfCore.addPageNumbers(mergedPdf, skipPageIndices);
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
      try {
        var attachments = nonInlineAttachments(item);
        var itemResults = [];

        var singleDoc = await PDFLib.PDFDocument.create();
        singleDoc.setProducer("Postfach2PDF (schantastic)");
        singleDoc.setCreator("Postfach2PDF");
        singleDoc.setTitle(item.subject || "E-Mail");
        singleDoc.setCreationDate(new Date());
        var skipPageIndices = [];

        await renderEmailPagesInto(item, options, attachments, singleDoc, itemResults, skipPageIndices, function (attachment, j, total) {
          setStatus(
            "E-Mail " + (i + 1) + "/" + messages.length + " – Anhang " + (j + 1) + "/" + total + ": " + attachment.name + " ..."
          );
        });

        if (options.pageNumbers) {
          await Postfach2PdfCore.addPageNumbers(singleDoc, skipPageIndices);
        }

        var bytes = await singleDoc.save();
        var fileName = buildFileName(item);
        Postfach2PdfCore.downloadBytes(bytes, fileName);

        overallResults.push({
          name: messages[i].subject || "(kein Betreff)",
          status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
          reason: "als eigene PDF (" + fileName + ") gespeichert",
        });
      } finally {
        await unloadItemAsync(item);
      }

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
