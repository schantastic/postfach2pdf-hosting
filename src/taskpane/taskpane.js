/* global Office, PDFLib, html2pdf, Postfach2PdfCore, Postfach2PdfI18n, document, window */

(function () {
  "use strict";

  var el = {};
  var currentAttachments = [];
  var batchSelection = [];

  var SETTINGS_KEYS = {
    loadExternalImages: "mailpdf.loadExternalImages",
    includeHeaders: "mailpdf.includeHeaders",
    pageNumbers: "mailpdf.pageNumbers",
    language: "mailpdf.language",
  };

  var INTL_LOCALE_TAGS = { de: "de-DE", en: "en-US", fr: "fr-FR", es: "es-ES", it: "it-IT" };

  function t(key, params) {
    return Postfach2PdfI18n.t(key, params);
  }

  // Bereits in der aktiven UI-Sprache aufgeloeste Textbausteine fuer
  // postfach2pdf-core.js (siehe DEFAULT_STRINGS dort fuer die erwarteten
  // Keys) - core.js selbst kennt i18n.js bewusst nicht (siehe
  // ARCHITECTURE.md), bekommt die Strings stattdessen hier fertig gebaut.
  function coreStrings() {
    return {
      placeholderHeading: t("pdfPlaceholderHeading"),
      labelFileName: t("pdfLabelFileName"),
      labelType: t("pdfLabelType"),
      labelSize: t("pdfLabelSize"),
      footerPageOfTemplate: t("pdfFooterPageOf"),
      footerBrand: "Postfach2PDF",
      unknownValue: t("unknownValue"),
      unknownSize: t("unknownSize"),
      reasonCloudAttachment: t("reasonCloudAttachment"),
      reasonEmbeddedOutlookItem: t("reasonEmbeddedOutlookItem"),
      reasonSizeLimitExceededTemplate: t("reasonSizeLimitExceeded"),
      reasonUnsupportedContentFormatTemplate: t("reasonUnsupportedContentFormat"),
      reasonPdfEmbedFailedTemplate: t("reasonPdfEmbedFailed"),
      reasonImageEmbedFailedTemplate: t("reasonImageEmbedFailed"),
      reasonUnsupportedFileTypeTemplate: t("reasonUnsupportedFileType"),
      reasonReadFailedTemplate: t("reasonReadFailed"),
    };
  }

  function getStoredLanguage() {
    var settings = Office.context.roamingSettings;
    if (!settings) {
      return "auto";
    }
    return settings.get(SETTINGS_KEYS.language) || "auto";
  }

  // Merkt sich, welcher der beiden "nicht unterstuetzt"-Texte zuletzt
  // gesetzt wurde (die werden per JS gesetzt, nicht per data-i18n, siehe
  // Office.onReady unten) - noetig, damit refreshDynamicContent() beim
  // Sprachwechsel den richtigen Text neu uebersetzen kann.
  var unsupportedReasonKey = null;

  // Alles, was NICHT ueber data-i18n automatisch von
  // applyStaticTranslations() erfasst wird (dynamisch aus Office.js-Daten
  // gebaute Listen/Texte), muss nach einem Sprachwechsel manuell neu
  // gerendert werden.
  function refreshDynamicContent() {
    if (unsupportedReasonKey) {
      el.unsupported.textContent = t(unsupportedReasonKey);
    }
    if (!el.app.hidden) {
      renderSummary();
      renderAttachmentList();
    }
    if (!el.batch.hidden) {
      renderBatchList();
    }
  }

  // Kein window.location.reload() (mehr): ein voller Neuladen der Taskpane
  // fuehrte in echtem Outlook dazu, dass die Sprachauswahl beim Neustart
  // wieder auf "Automatisch" zurueckfiel - vermutlich eine Race Condition
  // zwischen roamingSettings.saveAsync() und dem sofortigen Reload (analog
  // zu den bereits dokumentierten Timing-Problemen in der echten, im
  // Vergleich zu einer isolierten Testumgebung viel schwereren
  // Outlook-Webseite, siehe STATUS.md). Stattdessen wird die Sprache direkt
  // im laufenden Taskpane umgeschaltet, ganz ohne Neuladen.
  async function onLanguageChanged() {
    var lang = el.language.value;
    await Postfach2PdfI18n.init(lang === "auto" ? null : lang, "../lib/i18n/");
    document.documentElement.lang = Postfach2PdfI18n.getActiveLanguage();
    Postfach2PdfI18n.applyStaticTranslations(document);
    refreshDynamicContent();
    persistSetting(SETTINGS_KEYS.language, lang);
  }

  Office.onReady(async function (info) {
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
    el.language = document.getElementById("mp-language");

    el.batchList = document.getElementById("mp-batch-list");
    el.batchEmpty = document.getElementById("mp-batch-empty");
    el.batchMerge = document.getElementById("mp-batch-merge");
    el.batchIncludeHeaders = document.getElementById("mp-batch-include-headers");
    el.batchPageNumbers = document.getElementById("mp-batch-page-numbers");
    el.batchStart = document.getElementById("mp-batch-start");

    var storedLanguage = getStoredLanguage();
    await Postfach2PdfI18n.init(storedLanguage === "auto" ? null : storedLanguage, "../lib/i18n/");
    document.documentElement.lang = Postfach2PdfI18n.getActiveLanguage();
    Postfach2PdfI18n.applyStaticTranslations(document);
    el.language.value = storedLanguage;
    el.language.addEventListener("change", onLanguageChanged);

    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      unsupportedReasonKey = "errorUnsupportedMailbox18";
      el.unsupported.hidden = false;
      el.unsupported.textContent = t(unsupportedReasonKey);
      return;
    }

    el.shared.hidden = false;

    if (currentItem()) {
      initSingleItemMode();
    } else if (Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
      initBatchMode();
    } else {
      unsupportedReasonKey = "errorUnsupportedNoSelection13";
      el.unsupported.hidden = false;
      el.unsupported.textContent = t(unsupportedReasonKey);
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
      var locale = INTL_LOCALE_TAGS[Postfach2PdfI18n.getActiveLanguage()] || "en-US";
      return new Intl.DateTimeFormat(locale, {
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
        label.textContent = entry.name + " — " + (entry.detail || t("resultDetailEmbeddedDefault"));
      } else if (entry.status === "skipped") {
        icon.textContent = "–";
        icon.setAttribute("data-kind", "skip");
        label.textContent = entry.name + " — " + (entry.detail || t("resultDetailSkippedDefault"));
      } else {
        icon.textContent = "⚠";
        icon.setAttribute("data-kind", "warn");
        label.textContent = entry.name + " — " + (entry.reason || entry.detail || t("resultDetailErrorDefault"));
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
    el.subject.textContent = item.subject || t("fallbackNoSubject");
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
      nameSpan.textContent = attachment.name || t("fallbackUnnamed");

      var detailSpan = document.createElement("span");
      detailSpan.className = "mp-attachment-detail";
      detailSpan.textContent =
        (attachment.contentType || t("fallbackUnknownType")) +
        " · " +
        Postfach2PdfCore.formatBytes(attachment.size, t("unknownSize"));

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
      : t("fileNameUnknownDate");
    var senderPart =
      Postfach2PdfCore.sanitizeFileNamePart(
        (item.from && (item.from.displayName || item.from.emailAddress)) || t("fileNameUnknownSender")
      ) || t("fileNameUnknownSender");
    var subjectPart =
      Postfach2PdfCore.sanitizeFileNamePart(item.subject || t("fileNameNoSubject")) || t("fileNameNoSubject");
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
      [t("labelSubject"), escapeHtml(item.subject || t("fallbackNoSubject"))],
      [t("labelFrom"), escapeHtml(formatAddress(item.from))],
      [t("labelTo"), escapeHtml(formatAddressList(item.to))],
    ];
    if (item.cc && item.cc.length > 0) {
      headerRows.push([t("labelCc"), escapeHtml(formatAddressList(item.cc))]);
    }
    headerRows.push([t("labelDate"), escapeHtml(formatDate(item.dateTimeCreated))]);

    if (options.includeHeaders) {
      if (item.internetMessageId) {
        headerRows.push([t("labelMessageId"), escapeHtml(item.internetMessageId)]);
      }
      if (item.conversationId) {
        headerRows.push([t("labelConversationId"), escapeHtml(item.conversationId)]);
      }
    }

    if (selectedAttachments.length > 0) {
      headerRows.push([
        t("labelAttachments"),
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
      escapeHtml(item.subject || t("fallbackNoSubject")) +
      "</h1>" +
      '<table class="mp-pdf-header">' +
      headerHtml +
      "</table>" +
      '<hr class="mp-pdf-divider" />' +
      '<div class="mp-pdf-body">' +
      processedBody +
      "</div>" +
      // Unsichtbarer Puffer ganz am Ende (unbestaetigte Mitigation):
      // Nutzer meldet Content-Verlust am allerletzten Stueck grosser
      // Mails, unabhaengig von der genauen Seitenaufteilung - deutet
      // darauf hin, dass html2canvas den untersten Bereich der Leinwand
      // nicht mehr fertig zeichnet (vermutlich Ressourcen-/Zeitdruck in
      // der echten, schweren Outlook-Webseite). Dieser Puffer verschiebt
      // den eigentlichen Inhalt (inkl. Fusszeile) weg vom gefaehrdeten
      // aeussersten Rand - geht selbst nichts verloren, wenn er
      // abgeschnitten wird.
      '<div aria-hidden="true" style="height: 250px;"></div>' +
      "</div>"
    );
  }

  // !important auf den Positionierungs-Eigenschaften unseres eigenen
  // Headers (Betreff/Von/An/Datum): der rohe Mail-Body wird direkt danach
  // ins selbe DOM eingefuegt, inklusive dessen eigener <style>-Bloecke.
  // Viele Newsletter-Vorlagen nutzen unscoped Selektoren (z. B. einfach
  // "h1 { margin-top: 40px !important; }"), die dann versehentlich auch
  // unseren eigenen Betreff-<h1> treffen und ihn seitenmittig statt oben
  // beginnen lassen. !important stellt sicher, dass unser eigener Header
  // davon unbeeinflusst bleibt.
  var PDF_DOCUMENT_STYLE =
    "<style>" +
    ".mp-pdf-document { font-family: Calibri, Arial, sans-serif; font-size: 12px; color: #000; margin: 0 !important; padding: 0 !important; }" +
    ".mp-pdf-subject { font-size: 18px !important; margin: 0 0 10px 0 !important; padding: 0 !important; }" +
    ".mp-pdf-header { margin: 0 !important; }" +
    ".mp-pdf-header td { padding: 2px 8px 2px 0 !important; vertical-align: top; margin: 0 !important; }" +
    ".mp-pdf-label { color: #555; white-space: nowrap; font-weight: 600; }" +
    ".mp-pdf-divider { border: none; border-top: 1px solid #999; margin: 10px 0 !important; padding: 0 !important; }" +
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
    if (inlineAttachments.length === 0) {
      return;
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
          imgs.forEach(function (img) {
            var src = img.getAttribute("src") || "";
            var matchesCid = attachment.contentId && src === "cid:" + attachment.contentId;
            var matchesId = src.indexOf(attachment.id) !== -1;
            if (matchesCid || matchesId) {
              img.src = dataUrl;
            }
          });
        } catch (error) {
          console.warn("Postfach2PDF: Inline-Bild konnte nicht eingebettet werden", attachment.name, error);
        }
      })
    );
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

    await Promise.all(
      imgs.map(function (img) {
        var src = img.getAttribute("src") || "";
        if (!/^https?:/i.test(src)) {
          return Promise.resolve();
        }
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
          })
          .catch(function (error) {
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
    await embedInlineAttachmentImages(item, el.renderRoot);
    await inlineExternalImages(el.renderRoot);
    await waitForImages(el.renderRoot);
    // Kurze Pause, bevor html2canvas zu erfassen beginnt: bei grossen
    // Mails ging vereinzelt Inhalt ganz am Ende verloren, vermutlich durch
    // Zeit-/Ressourcendruck in der (im Vergleich zu unserer Taskpane viel
    // schwereren) echten Outlook-Webseite. Gibt Layout/GC etwas Luft, sich
    // vor der aufwendigen Erfassung zu setzen.
    await new Promise(function (resolve) {
      setTimeout(resolve, 300);
    });
    // Fusszeile/Seitenzahl (addPageNumbers) braucht ~30pt Platz am unteren
    // Rand jeder Seite. Bei 15pt Rand ueberlappt der letzte Textabschnitt
    // einer vollen Seite die Fusszeile (mit echten Chromium-Renderings
    // reproduziert und verifiziert) - deshalb mehr Rand, wenn Seitenzahlen
    // aktiv sind.
    var bottomMargin = options.pageNumbers ? 45 : 15;
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
            // Reduziert Speicherbedarf der Erfassungsflaeche bei grossen
            // Mails (siehe Pause oben) - auf Kosten leicht geringerer
            // Bildschaerfe.
            scale: 1.5,
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
      return arrayBuffer;
    } finally {
      el.renderRoot.innerHTML = "";
    }
  }

  // Rendert eine einzelne E-Mail (Body + alle Anhaenge) als Seiten in das
  // uebergebene PDFDocument. Wird sowohl vom Einzel- als auch vom
  // Batch-Modus genutzt. skipPageIndices sammelt die 0-basierten Indizes
  // seitengetreu kopierter PDF-Anhangsseiten (kein reservierter Rand) -
  // addPageNumbers() darf darauf nichts zeichnen, siehe dort.
  async function renderEmailPagesInto(item, options, attachments, targetPdf, results, skipPageIndices, onProgress) {
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
      var pageCountBefore = targetPdf.getPageCount();
      var outcome = await Postfach2PdfCore.embedAttachmentIntoPdf(item, targetPdf, attachment, coreStrings());
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
      setStatus(t("statusProcessingAttachment", { index: i + 1, total: total, name: attachment.name }));
    });

    deselectedAttachments.forEach(function (attachment) {
      results.push({ name: attachment.name, status: "skipped" });
    });

    if (options.pageNumbers) {
      setStatus(t("statusAddingPageNumbers"));
      await Postfach2PdfCore.addPageNumbers(mergedPdf, skipPageIndices, coreStrings());
    }

    setStatus(t("statusCreatingPdfBytes"));
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
    setStatus(t("statusGeneratingPdf"));

    var item = currentItem();

    try {
      var built = await buildFinalPdf(item, currentOptions());
      setStatus(t("statusSavingPdf"));
      Postfach2PdfCore.downloadBytes(built.bytes, built.fileName);
      renderResults(built.results);
      setStatus(t("statusPdfSuccess"), "success");
    } catch (error) {
      console.error("Postfach2PDF: PDF-Erzeugung fehlgeschlagen", error);
      setStatus(
        t("errorPdfGenerationPrefix", { message: error && error.message ? error.message : error }),
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
      nameSpan.textContent = message.subject || t("fallbackNoSubject");

      var detailSpan = document.createElement("span");
      detailSpan.className = "mp-attachment-detail";
      detailSpan.textContent = message.hasAttachments ? t("fallbackHasAttachments") : t("fallbackNoAttachmentsShort");

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
      new Date().toISOString().slice(0, 10) + "_Postfach2PDF-Batch_" + messages.length + "-" + t("fileNameMailsUnit")
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
      setStatus(t("statusLoadingEmail", { index: i + 1, total: messages.length, subject: messages[i].subject || "" }));
      var item = await loadItemByIdAsync(messages[i].itemId);
      try {
        var attachments = nonInlineAttachments(item);
        var itemResults = [];
        await renderEmailPagesInto(item, options, attachments, mergedPdf, itemResults, skipPageIndices, function (attachment, j, total) {
          setStatus(
            t("statusEmailAttachmentProgress", {
              emailIndex: i + 1,
              emailTotal: messages.length,
              attIndex: j + 1,
              attTotal: total,
              name: attachment.name,
            })
          );
        });
        overallResults.push({
          name: messages[i].subject || t("fallbackNoSubject"),
          status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
          reason: t("resultMergedIntoSharedPdf", { count: itemResults.length }),
        });
      } finally {
        await unloadItemAsync(item);
      }
    }

    if (options.pageNumbers) {
      setStatus(t("statusAddingPageNumbers"));
      await Postfach2PdfCore.addPageNumbers(mergedPdf, skipPageIndices, coreStrings());
    }

    var bytes = await mergedPdf.save();
    var fileName = buildBatchFileName(messages);
    Postfach2PdfCore.downloadBytes(bytes, fileName);
    return overallResults;
  }

  async function convertBatchSeparate(messages, options) {
    var overallResults = [];

    for (var i = 0; i < messages.length; i++) {
      setStatus(t("statusLoadingEmail", { index: i + 1, total: messages.length, subject: messages[i].subject || "" }));
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
            t("statusEmailAttachmentProgress", {
              emailIndex: i + 1,
              emailTotal: messages.length,
              attIndex: j + 1,
              attTotal: total,
              name: attachment.name,
            })
          );
        });

        if (options.pageNumbers) {
          await Postfach2PdfCore.addPageNumbers(singleDoc, skipPageIndices, coreStrings());
        }

        var bytes = await singleDoc.save();
        var fileName = buildFileName(item);
        Postfach2PdfCore.downloadBytes(bytes, fileName);

        overallResults.push({
          name: messages[i].subject || t("fallbackNoSubject"),
          status: itemResults.some(function (r) { return r.status !== "embedded"; }) ? "failed" : "embedded",
          reason: t("resultSavedAsSeparatePdf", { fileName: fileName }),
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
    setStatus(t("statusStartingBatch"));

    var messages = getSelectedBatchMessages();
    if (messages.length === 0) {
      setStatus(t("errorNoEmailsSelected"), "error");
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
        setStatus(Postfach2PdfI18n.plural("batchSuccess", messages.length), "success");
      } else {
        setStatus(t("statusBatchPartialFailure", { failed: failedCount, total: messages.length }), "error");
      }
    } catch (error) {
      console.error("Postfach2PDF: Batch-Verarbeitung fehlgeschlagen", error);
      setStatus(
        t("errorBatchPrefix", { message: error && error.message ? error.message : error }),
        "error"
      );
    } finally {
      el.batchStart.disabled = false;
    }
  }
})();
