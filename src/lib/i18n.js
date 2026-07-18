/* global Office, fetch, document, window */
/**
 * Postfach2PDF i18n: laedt Uebersetzungs-JSONs, erkennt die Sprache ueber
 * Office.context.displayLanguage, stellt t()/applyStaticTranslations()
 * bereit. Haengt sich als window.Postfach2PdfI18n an. Wird bewusst nicht
 * von postfach2pdf-core.js geladen (siehe ARCHITECTURE.md) - core.js
 * bekommt bereits aufgeloeste Strings als Parameter uebergeben.
 */
(function (global) {
  "use strict";

  var SUPPORTED_LANGUAGES = ["de", "en", "fr", "es", "it"];
  var FALLBACK_LANGUAGE = "en";

  var strings = {};
  var fallbackStrings = {};
  var activeLanguage = FALLBACK_LANGUAGE;

  function detectLanguage() {
    var raw = "";
    try {
      raw = (Office.context && Office.context.displayLanguage) || "";
    } catch (e) {
      raw = "";
    }
    if (!raw && typeof navigator !== "undefined") {
      raw = navigator.language || "";
    }
    var prefix = String(raw).split("-")[0].toLowerCase();
    return SUPPORTED_LANGUAGES.indexOf(prefix) !== -1 ? prefix : FALLBACK_LANGUAGE;
  }

  function loadJson(path) {
    return fetch(path).then(function (response) {
      if (!response.ok) {
        throw new Error("HTTP " + response.status);
      }
      return response.json();
    });
  }

  // preferredLang: "auto" oder eine der SUPPORTED_LANGUAGES (z. B. aus
  // roamingSettings). basePath: Pfad zum i18n-Ordner relativ zur
  // aufrufenden HTML-Datei (taskpane.html und compose.html liegen in
  // unterschiedlichen Verzeichnissen, daher kein fester Pfad hier).
  function init(preferredLang, basePath) {
    var path = basePath || "../lib/i18n/";
    var lang =
      preferredLang && SUPPORTED_LANGUAGES.indexOf(preferredLang) !== -1 ? preferredLang : detectLanguage();
    activeLanguage = lang;

    var fallbackPromise =
      lang === FALLBACK_LANGUAGE
        ? Promise.resolve()
        : loadJson(path + FALLBACK_LANGUAGE + ".json")
            .then(function (json) {
              fallbackStrings = json;
            })
            .catch(function () {
              fallbackStrings = {};
            });

    return loadJson(path + lang + ".json")
      .then(function (json) {
        strings = json;
      })
      .catch(function (error) {
        console.warn("Postfach2PDF: Uebersetzung fuer '" + lang + "' konnte nicht geladen werden", error);
        strings = {};
      })
      .then(function () {
        return fallbackPromise;
      })
      .then(function () {
        return activeLanguage;
      });
  }

  function lookup(key) {
    if (Object.prototype.hasOwnProperty.call(strings, key)) {
      return strings[key];
    }
    if (Object.prototype.hasOwnProperty.call(fallbackStrings, key)) {
      return fallbackStrings[key];
    }
    // Sichtbarer Hinweis auf einen fehlenden Key statt leerem Text -
    // sollte im Normalbetrieb nie auftreten (de/en sind immer vollstaendig).
    return "[[" + key + "]]";
  }

  function interpolate(template, params) {
    if (!params) {
      return template;
    }
    return template.replace(/\{(\w+)\}/g, function (match, name) {
      return Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : match;
    });
  }

  function t(key, params) {
    return interpolate(lookup(key), params);
  }

  function pluralSuffix(count) {
    if (activeLanguage === "fr") {
      return count === 0 || count === 1 ? "_one" : "_other";
    }
    return count === 1 ? "_one" : "_other";
  }

  // Waehlt zwischen "{baseKey}_one"/"{baseKey}_other" je nach Sprache und
  // Anzahl, reicht "count" automatisch als Platzhalter durch.
  function plural(baseKey, count, params) {
    var key = baseKey + pluralSuffix(count);
    var mergedParams = { count: count };
    if (params) {
      for (var name in params) {
        if (Object.prototype.hasOwnProperty.call(params, name)) {
          mergedParams[name] = params[name];
        }
      }
    }
    return t(key, mergedParams);
  }

  // Attribute-Praefix fuer beliebige zu uebersetzende Attribute, z. B.
  // data-i18n-attr-aria-label="labelLanguage" setzt aria-label. Deckt
  // title/placeholder/aria-label etc. einheitlich ab statt je ein
  // Spezialfall-Attribut zu brauchen.
  var ATTR_PREFIX = "data-i18n-attr-";

  function applyStaticTranslations(rootEl) {
    var root = rootEl || document;
    var textNodes = root.querySelectorAll("[data-i18n]");
    for (var i = 0; i < textNodes.length; i++) {
      textNodes[i].textContent = t(textNodes[i].getAttribute("data-i18n"));
    }

    var allNodes = root.querySelectorAll("*");
    for (var n = 0; n < allNodes.length; n++) {
      var node = allNodes[n];
      for (var a = 0; a < node.attributes.length; a++) {
        var attr = node.attributes[a];
        if (attr.name.indexOf(ATTR_PREFIX) === 0) {
          var targetAttr = attr.name.slice(ATTR_PREFIX.length);
          node.setAttribute(targetAttr, t(attr.value));
        }
      }
    }
  }

  function getActiveLanguage() {
    return activeLanguage;
  }

  global.Postfach2PdfI18n = {
    SUPPORTED_LANGUAGES: SUPPORTED_LANGUAGES,
    FALLBACK_LANGUAGE: FALLBACK_LANGUAGE,
    detectLanguage: detectLanguage,
    init: init,
    t: t,
    plural: plural,
    applyStaticTranslations: applyStaticTranslations,
    getActiveLanguage: getActiveLanguage,
  };
})(window);
