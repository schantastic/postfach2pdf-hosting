/* global Office, PDFLib, Postfach2PdfCore, document, window */

(function () {
  "use strict";

  var el = {};
  var currentAttachments = [];

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      return;
    }

    el.app = document.getElementById("mp-app");
    el.unsupported = document.getElementById("mp-unsupported");
    el.attachmentList = document.getElementById("mp-attachment-list");
    el.noAttachments = document.getElementById("mp-no-attachments");
    el.merge = document.getElementById("mp-merge");
    el.removeOriginals = document.getElementById("mp-remove-originals");
    el.convert = document.getElementById("mp-convert");
    el.status = document.getElementById("mp-status");
    el.result = document.getElementById("mp-result");
    el.resultList = document.getElementById("mp-result-list");

    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      el.unsupported.hidden = false;
      return;
    }

    el.app.hidden = false;
    loadAttachments();
    el.convert.addEventListener("click", onConvertClicked);
  });

  function currentItem() {
    return Office.context.mailbox.item;
  }

  function getAttachmentsAsync() {
    return new Promise(function (resolve, reject) {
      currentItem().getAttachmentsAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function getSubjectAsync() {
    return new Promise(function (resolve) {
      var item = currentItem();
      if (item.subject && typeof item.subject.getAsync === "function") {
        item.subject.getAsync(function (result) {
          resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
        });
      } else {
        resolve(item.subject || "");
      }
    });
  }

  function addFileAttachmentFromBase64Async(base64, filename) {
    return new Promise(function (resolve, reject) {
      currentItem().addFileAttachmentFromBase64Async(base64, filename, { isInline: false }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function removeAttachmentAsync(attachmentId) {
    return new Promise(function (resolve, reject) {
      currentItem().removeAttachmentAsync(attachmentId, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      });
    });
  }

  function nonInlineAttachments(list) {
    return (list || []).filter(function (a) {
      return !a.isInline;
    });
  }

  async function loadAttachments() {
    try {
      var all = await getAttachmentsAsync();
      currentAttachments = nonInlineAttachments(all);
      renderAttachmentList();
    } catch (error) {
      console.error("Postfach2PDF: Anhaenge konnten nicht geladen werden", error);
      setStatus(
        "Anhaenge konnten nicht geladen werden: " + (error && error.message ? error.message : error),
        "error"
      );
    }
  }

  function renderAttachmentList() {
    el.attachmentList.innerHTML = "";
    if (currentAttachments.length === 0) {
      el.noAttachments.hidden = false;
      return;
    }
    el.noAttachments.hidden = true;

    currentAttachments.forEach(function (attachment, index) {
      var convertible = Postfach2PdfCore.isConvertibleAttachment(attachment);
      var li = document.createElement("li");
      li.className = "mp-attachment-item";

      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = convertible;
      checkbox.disabled = !convertible;
      checkbox.id = "mp-att-" + index;

      var label = document.createElement("label");
      label.className = "mp-attachment-meta";
      label.setAttribute("for", checkbox.id);

      var nameSpan = document.createElement("span");
      nameSpan.className = "mp-attachment-name";
      nameSpan.textContent = attachment.name || "(unbenannt)";

      var detailSpan = document.createElement("span");
      detailSpan.className = "mp-attachment-detail";
      detailSpan.textContent =
        (attachment.contentType || "unbekannter Typ") +
        " · " +
        Postfach2PdfCore.formatBytes(attachment.size) +
        (convertible ? "" : " · wird nicht unterstuetzt, bleibt unveraendert erhalten");

      label.appendChild(nameSpan);
      label.appendChild(detailSpan);
      li.appendChild(checkbox);
      li.appendChild(label);
      el.attachmentList.appendChild(li);
    });
  }

  function getSelectedConvertibleAttachments() {
    return currentAttachments.filter(function (attachment, index) {
      var checkbox = document.getElementById("mp-att-" + index);
      return checkbox && checkbox.checked && !checkbox.disabled;
    });
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
      } else if (entry.status === "skipped") {
        icon.textContent = "–";
        icon.setAttribute("data-kind", "skip");
      } else {
        icon.textContent = "⚠";
        icon.setAttribute("data-kind", "warn");
      }
      label.textContent = entry.name + " — " + entry.detail;

      li.appendChild(icon);
      li.appendChild(label);
      el.resultList.appendChild(li);
    });
    el.result.hidden = results.length === 0;
  }

  function buildMergedFileName(subject) {
    var subjectPart = Postfach2PdfCore.sanitizeFileNamePart(subject || "Anhaenge") || "Anhaenge";
    var stem = Postfach2PdfCore.sanitizeFileNameStem(subjectPart + "_Anhaenge");
    return stem + ".pdf";
  }

  function buildSingleFileName(attachmentName, usedNames) {
    var withoutExt = String(attachmentName || "Anhang").replace(/\.[^.]+$/, "");
    var base = Postfach2PdfCore.sanitizeFileNamePart(withoutExt) || "Anhang";
    var stem = Postfach2PdfCore.sanitizeFileNameStem(base);
    var candidate = stem + ".pdf";
    var counter = 2;
    while (usedNames.indexOf(candidate.toLowerCase()) !== -1) {
      candidate = stem + " (" + counter + ").pdf";
      counter++;
    }
    usedNames.push(candidate.toLowerCase());
    return candidate;
  }

  async function convertMerged(item, attachments, removeOriginals, results) {
    var mergedPdf = await PDFLib.PDFDocument.create();
    mergedPdf.setProducer("Postfach2PDF (schantastic)");
    mergedPdf.setCreator("Postfach2PDF");
    mergedPdf.setCreationDate(new Date());

    var embeddedOriginalIds = [];

    for (var i = 0; i < attachments.length; i++) {
      var attachment = attachments[i];
      setStatus("Verarbeite " + (i + 1) + "/" + attachments.length + ": " + attachment.name + " ...");
      var outcome = await Postfach2PdfCore.embedAttachmentIntoPdf(item, mergedPdf, attachment);
      if (outcome.status === "embedded") {
        embeddedOriginalIds.push(attachment.id);
        results.push({ name: attachment.name, status: "embedded", detail: "in gemeinsame PDF eingebettet" });
      } else {
        results.push({
          name: attachment.name,
          status: "failed",
          detail: outcome.reason + " (Original bleibt erhalten)",
        });
      }
    }

    if (embeddedOriginalIds.length === 0) {
      return;
    }

    setStatus("Fuege gemeinsame PDF an die Nachricht an ...");
    var subject = await getSubjectAsync();
    var fileName = buildMergedFileName(subject);
    var bytes = await mergedPdf.save();
    var base64 = Postfach2PdfCore.uint8ArrayToBase64(bytes);
    await addFileAttachmentFromBase64Async(base64, fileName);

    if (removeOriginals) {
      for (var j = 0; j < embeddedOriginalIds.length; j++) {
        try {
          await removeAttachmentAsync(embeddedOriginalIds[j]);
        } catch (error) {
          console.error("Postfach2PDF: Original-Anhang konnte nicht entfernt werden", error);
        }
      }
    }
  }

  async function convertSeparate(item, attachments, removeOriginals, results) {
    var usedNames = [];
    for (var i = 0; i < attachments.length; i++) {
      var attachment = attachments[i];
      setStatus("Verarbeite " + (i + 1) + "/" + attachments.length + ": " + attachment.name + " ...");
      var singleDoc = await PDFLib.PDFDocument.create();
      singleDoc.setProducer("Postfach2PDF (schantastic)");
      singleDoc.setCreator("Postfach2PDF");
      singleDoc.setCreationDate(new Date());

      var outcome = await Postfach2PdfCore.embedAttachmentIntoPdf(item, singleDoc, attachment);
      if (outcome.status !== "embedded") {
        results.push({
          name: attachment.name,
          status: "failed",
          detail: outcome.reason + " (Original bleibt erhalten)",
        });
        continue;
      }

      var fileName = buildSingleFileName(attachment.name, usedNames);
      var bytes = await singleDoc.save();
      var base64 = Postfach2PdfCore.uint8ArrayToBase64(bytes);
      await addFileAttachmentFromBase64Async(base64, fileName);

      if (removeOriginals) {
        try {
          await removeAttachmentAsync(attachment.id);
        } catch (error) {
          console.error("Postfach2PDF: Original-Anhang konnte nicht entfernt werden", error);
        }
      }
      results.push({ name: attachment.name, status: "embedded", detail: "als " + fileName + " angehaengt" });
    }
  }

  async function onConvertClicked() {
    el.convert.disabled = true;
    el.result.hidden = true;
    setStatus("Starte Konvertierung ...");

    var item = currentItem();
    var selected = getSelectedConvertibleAttachments();
    var merge = el.merge.checked;
    var removeOriginals = el.removeOriginals.checked;
    var results = [];

    currentAttachments
      .filter(function (attachment) {
        return selected.indexOf(attachment) === -1;
      })
      .forEach(function (attachment) {
        results.push({
          name: attachment.name,
          status: "skipped",
          detail: Postfach2PdfCore.isConvertibleAttachment(attachment)
            ? "nicht ausgewaehlt"
            : "Dateityp wird in dieser Version nicht unterstuetzt",
        });
      });

    if (selected.length === 0) {
      renderResults(results);
      setStatus("Keine konvertierbaren Anhaenge ausgewaehlt.", "error");
      el.convert.disabled = false;
      return;
    }

    try {
      if (merge) {
        await convertMerged(item, selected, removeOriginals, results);
      } else {
        await convertSeparate(item, selected, removeOriginals, results);
      }

      renderResults(results);
      var failedCount = results.filter(function (r) {
        return r.status === "failed";
      }).length;
      if (failedCount === 0) {
        setStatus("Anhaenge wurden konvertiert.", "success");
      } else {
        setStatus(failedCount + " Anhang/Anhaenge konnten nicht konvertiert werden, siehe Ergebnis.", "error");
      }
      await loadAttachments();
    } catch (error) {
      console.error("Postfach2PDF: Konvertierung fehlgeschlagen", error);
      setStatus("Fehler bei der Konvertierung: " + (error && error.message ? error.message : error), "error");
    } finally {
      el.convert.disabled = false;
    }
  }
})();
