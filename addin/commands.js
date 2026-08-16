/* global Office */

// ---------------------------------------------------------------------------
// Auto-Antwort: liest den Volltext der aktuellen Nachricht und öffnet den
// Standardbrowser mit https://claude.ai/new?q=<Prompt>. Es wird NICHTS
// automatisch gesendet – der Nutzer klickt auf claude.ai selbst auf Submit.
// Läuft komplett clientseitig; Mailinhalte erreichen nie den Hosting-Server.
// ---------------------------------------------------------------------------

// Instruktion, die dem Mailtext vorangestellt wird (frei anpassbar):
var PROMPT_PREFIX =
  "Bitte entwirf eine passende Antwort auf die folgende E-Mail entsprechend meines Email-Skills. " +
  "Antworte in der Sprache der E-Mail. Beachte alternativ exakt meine folgende Anweisung nach dem Hauptemailtext.\n\n---\n\n";

// Wird IMMER ans Ende angehaengt (auch bei gekuerzten Mails) - der Cursor auf
// claude.ai landet darunter, dort kann direkt eine eigene Anweisung getippt werden:
var PROMPT_SUFFIX = "\n\n----------\n\n";

// Sicherheitslimit für die URL-Länge (Zeichen NACH dem URL-Encoding).
// Browser/Server werden ab ~8.000 Zeichen unzuverlässig; 6.000 ist konservativ.
var MAX_ENCODED_LEN = 6000;

var CLAUDE_URL = "https://claude.ai/new?q=";

Office.onReady(function () {});

/**
 * Ribbon-Kommando (ExecuteFunction). Funktioniert auf der Read-Oberfläche
 * (geöffnete Nachricht) und der Compose-Oberfläche (inkl. Inline-Antwort).
 */
function autoAntwort(event) {
  var item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notify(item, "Auto-Antwort: Nachrichtentext konnte nicht gelesen werden.");
      event.completed();
      return;
    }

    buildHeader(item, function (header) {
      var prompt = PROMPT_PREFIX + header + (result.value || "").trim();
      openClaude(prompt);
      notify(item, "Auto-Antwort: claude.ai wurde im Browser geöffnet – dort prüfen und absenden.");
      event.completed();
    });
  });
}

/**
 * Betreff/Absender ermitteln. Im Read-Modus sind das einfache Strings,
 * im Compose-Modus ist subject ein Objekt mit getAsync.
 */
function buildHeader(item, callback) {
  var from = "";
  if (item.from && item.from.displayName) {
    from = "Von: " + item.from.displayName + " <" + item.from.emailAddress + ">\n";
  }

  if (typeof item.subject === "string") {
    callback("Betreff: " + item.subject + "\n" + from + "\n");
  } else if (item.subject && typeof item.subject.getAsync === "function") {
    item.subject.getAsync(function (res) {
      var subject = res.status === Office.AsyncResultStatus.Succeeded ? res.value : "";
      callback(subject ? "Betreff: " + subject + "\n\n" : "");
    });
  } else {
    callback("");
  }
}

/**
 * URL bauen, auf sichere Länge kürzen und im Standardbrowser öffnen.
 * window.open mit externer URL öffnet aus Outlook-Add-in-Kommandos heraus
 * den Standardbrowser des Systems (nicht die Claude-App).
 */
function openClaude(prompt) {
  // Platz für den Suffix (Trennlinie) reservieren, damit er nie der Kürzung
  // zum Opfer fällt und der Cursor immer unterhalb der Linie landet.
  var encodedSuffix = encodeURIComponent(PROMPT_SUFFIX);
  var maxBody = MAX_ENCODED_LEN - encodedSuffix.length;
  var encoded = encodeURIComponent(prompt);

  if (encoded.length > maxBody) {
    // Rohtext schrittweise kürzen, bis die kodierte Länge passt.
    var raw = prompt;
    while (encoded.length > maxBody && raw.length > 0) {
      var over = encoded.length - maxBody;
      raw = raw.slice(0, raw.length - Math.max(50, Math.ceil(over / 3)));
      encoded = encodeURIComponent(raw + "\n\n[Text wurde wegen Längenbegrenzung gekürzt]");
    }
  }

  window.open(CLAUDE_URL + encoded + encodedSuffix);
}

/** Dezente Infoleiste in der Nachricht (kein blockierender Dialog). */
function notify(item, text) {
  try {
    item.notificationMessages.replaceAsync("autoAntwortInfo", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: text.substring(0, 150),
      icon: "Icon.16",
      persistent: false
    });
  } catch (e) { /* Benachrichtigung ist optional */ }
}

// Funktion für ExecuteFunction registrieren (erforderlich für klassisches Outlook)
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("autoAntwort", autoAntwort);
}
// Ältere Laufzeiten suchen die Funktion am globalen Objekt:
window.autoAntwort = autoAntwort;
