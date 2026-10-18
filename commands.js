/* Werbe-Widerspruch: Outlook-Add-in (neues Outlook, Outlook im Web, Mac)
 *
 * Schaltfläche "Widerspruch" in einer geöffneten Werbemail:
 *  - prüft, ob es Spam ist (dann keine Antwort, sonst wird die Adresse bestätigt)
 *  - sucht eine erreichbare Adresse (Reply-To, Absender, Datenschutz/Impressum, info@)
 *  - öffnet einen fertigen Widerspruch als Entwurf, du musst nur noch senden
 */

// ===================== EINSTELLUNGEN =====================
const EINSTELLUNGEN = {
  // leer = Anzeigename aus deinem Postfach
  meinName: "",
  // Ohne Abmelde-Kopf (List-Unsubscribe) keine Antwort: seriöse Newsletter haben ihn immer.
  nurMitAbmeldeKopf: true,
  // Ab dieser Spam-Bewertung von Exchange (SCL 0-9) keine Antwort
  spamAbScl: 5,
  // Sperrfrist pro Absender, danach verschärfte zweite Mail
  erneutNachTagen: 14,
  mitAuskunft: true,
  fallbackInfo: true,
  abmeldeAdresseInCc: true,
  kategorie: "Widerspruch versandt",
};
// =========================================================

const NOREPLY = ["noreply", "donotreply", "dontreply", "mailerdaemon", "bounce", "postmaster", "nichtantworten", "keineantwort"];
const PRAEFIXE = ["datenschutz", "privacy", "dataprotection", "data-protection", "dsb", "dpo", "gdpr", "dsgvo",
  "impressum", "legal", "recht", "info", "kontakt", "contact", "service", "kundenservice", "support", "hallo", "hello"];

function asyncCall(fn) {
  return new Promise((resolve, reject) => {
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
  });
}

function istNoReply(adr) {
  const lokal = (adr.split("@")[0] || "").toLowerCase().replace(/[-_.]/g, "");
  return NOREPLY.some((m) => lokal.includes(m));
}

function basisDomain(adr) {
  if (!adr || !adr.includes("@")) return "";
  const t = adr.toLowerCase().split("@").pop().split(".");
  if (t.length < 2) return "";
  let b = t.slice(-2).join(".");
  if (t.length >= 3 && (t[t.length - 2] === "co" || t[t.length - 2] === "com")) b = t.slice(-3).join(".");
  return b;
}

function adressenImText(s) {
  const treffer = (s || "").match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi) || [];
  return [...new Set(treffer.map((a) => a.toLowerCase().replace(/^\.+/, "").replace(/[.-]+$/, "")))];
}

// Kopfzeilen in {feld: wert} (Folgezeilen werden angehängt, Feldnamen klein)
function kopfZerlegen(roh) {
  const k = {};
  let letztes = null;
  for (const zeile of (roh || "").split(/\r?\n/)) {
    if (/^[ \t]/.test(zeile) && letztes) {
      k[letztes] += " " + zeile.trim();
    } else {
      const p = zeile.indexOf(":");
      if (p > 0) {
        letztes = zeile.slice(0, p).trim().toLowerCase();
        k[letztes] = (k[letztes] ? k[letztes] + ", " : "") + zeile.slice(p + 1).trim();
      }
    }
  }
  return k;
}

function spamWert(k) {
  const scl = k["x-ms-exchange-organization-scl"];
  if (scl && !isNaN(parseInt(scl, 10))) return parseInt(scl, 10);
  const f = k["x-forefront-antispam-report"] || "";
  if (/SFV:SPM/i.test(f)) return 9;
  const m = f.match(/SCL:(-?\d+)/i);
  return m ? parseInt(m[1], 10) : -1;
}

function abmeldeAdresse(wert) {
  const m = (wert || "").match(/mailto:([^>?,\s]+)/i);
  return m ? decodeURIComponent(m[1]).toLowerCase() : "";
}

function besteAdresseAusMail(text, absender, meine) {
  let bester = 999, kandidat = "", gleicheDomain = "";
  for (const a of adressenImText(text)) {
    if (istNoReply(a) || basisDomain(a) === basisDomain(meine)) continue;
    const rang = PRAEFIXE.findIndex((p) => a.startsWith(p));
    if (rang >= 0 && rang < bester) { bester = rang; kandidat = a; }
    if (!gleicheDomain && basisDomain(a) === basisDomain(absender)) gleicheDomain = a;
  }
  return kandidat || gleicheDomain;
}

function datum(d) {
  return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function brieftext(item, meineAdr, frueher, name) {
  const absaetze = [];
  absaetze.push(esc("Sehr geehrte Damen und Herren,"));
  absaetze.push(esc(`am ${datum(item.dateTimeCreated)} haben Sie an meine E-Mail-Adresse ${meineAdr} eine Werbe-E-Mail mit dem Betreff „${item.subject}“ geschickt (Absender: ${item.from.emailAddress}).`));
  if (frueher) {
    absaetze.push(esc(`Bereits mit E-Mail vom ${datum(frueher)} habe ich Ihnen mitgeteilt, dass ich keine Werbung von Ihnen wünsche, und Sie zur Unterlassung aufgefordert. Dass Sie mir trotzdem erneut Werbung schicken, ist ein bewusster Verstoß, den ich nicht weiter hinnehmen werde.`));
  }
  absaetze.push(esc("Ich habe Ihren Newsletter nicht bestellt und in den Erhalt von Werbung per E-Mail zu keinem Zeitpunkt ausdrücklich eingewilligt. Die Zusendung ist eine unzumutbare Belästigung im Sinne von § 7 Abs. 2 Nr. 2 UWG und ein rechtswidriger Eingriff in mein allgemeines Persönlichkeitsrecht (§§ 823 Abs. 1, 1004 Abs. 1 BGB analog). Einer etwaigen Verwendung meiner Adresse nach § 7 Abs. 3 UWG widerspreche ich hiermit ausdrücklich."));
  const punkte = [
    "die Zusendung jeglicher Werbung an mich, gleich an welche meiner E-Mail-Adressen, ab sofort zu unterlassen;",
    "meine personenbezogenen Daten zu löschen (Art. 17 DSGVO). Der Verarbeitung meiner Daten zu Werbezwecken widerspreche ich nach Art. 21 Abs. 2 DSGVO; sie dürfen dafür nicht mehr verarbeitet werden (Art. 21 Abs. 3 DSGVO);",
  ];
  if (EINSTELLUNGEN.mitAuskunft) {
    punkte.push("mir nach Art. 15 DSGVO mitzuteilen, welche Daten Sie über mich gespeichert haben, woher Sie meine Adresse haben und an wen Sie sie weitergegeben haben. Berufen Sie sich auf eine Einwilligung, legen Sie bitte deren Nachweis vor (Art. 7 Abs. 1 DSGVO);");
  }
  punkte.push("mir die Erledigung binnen 14 Tagen schriftlich zu bestätigen.");
  absaetze.push("Ich fordere Sie auf,<br>" + punkte.map((p, i) => `${i + 1}. ${esc(p)}`).join("<br>"));
  if (frueher) {
    absaetze.push(esc("Sollte ich innerhalb dieser Frist keine Bestätigung erhalten oder erneut Werbung von Ihnen bekommen, werde ich ohne weitere Ankündigung anwaltlich abmahnen lassen und Beschwerde bei der zuständigen Datenschutzaufsichtsbehörde einlegen. Die Kosten dafür haben Sie zu tragen."));
  } else {
    absaetze.push(esc("Bei jeder weiteren unerwünschten Zusendung behalte ich mir rechtliche Schritte, insbesondere eine kostenpflichtige Abmahnung und eine Beschwerde bei der zuständigen Datenschutzaufsichtsbehörde, ausdrücklich vor."));
  }
  absaetze.push("Mit freundlichen Grüßen<br>" + esc(name));
  return absaetze.map((a) => `<p>${a}</p>`).join("");
}

// ---------- Protokoll (wandert mit dem Postfach, auch auf andere Geräte) ----------
function protokollLesen() {
  return Office.context.roamingSettings.get("protokoll") || {};
}
async function protokollSchreiben(absender) {
  const p = protokollLesen();
  const grenze = Date.now() - 400 * 864e5;
  for (const a of Object.keys(p)) if (new Date(p[a]).getTime() < grenze) delete p[a];
  p[absender] = new Date().toISOString().slice(0, 10);
  Office.context.roamingSettings.set("protokoll", p);
  await asyncCall((cb) => Office.context.roamingSettings.saveAsync(cb));
}

async function kategorieSetzen(item) {
  const name = EINSTELLUNGEN.kategorie;
  try {
    const vorhanden = await asyncCall((cb) => Office.context.mailbox.masterCategories.getAsync(cb));
    if (!(vorhanden || []).some((c) => c.displayName === name)) {
      await asyncCall((cb) => Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: name, color: Office.MailboxEnums.CategoryColor.Preset0 }], cb));
    }
    await asyncCall((cb) => item.categories.addAsync([name], cb));
  } catch (e) { /* Kategorie ist nur Komfort */ }
}

function melden(text, fehler) {
  const item = Office.context.mailbox.item;
  const msg = fehler
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: text.slice(0, 150) }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: text.slice(0, 150),
        icon: "Icon.16", persistent: false };
  item.notificationMessages.replaceAsync("widerspruch", msg);
}

// ---------- Hauptablauf ----------
async function pruefenUndErstellen() {
  const item = Office.context.mailbox.item;
  const meine0 = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  const absender = (item.from && item.from.emailAddress || "").toLowerCase();
  if (!absender.includes("@")) return "Keine Absenderadresse gefunden.";
  if (basisDomain(absender) === basisDomain(meine0)) return "Interner Absender, kein Widerspruch.";

  const k = kopfZerlegen(await asyncCall((cb) => item.getAllInternetHeadersAsync(cb)));
  const scl = spamWert(k);
  if (scl >= EINSTELLUNGEN.spamAbScl) return `Sieht nach Spam aus (SCL ${scl}). Keine Antwort, sonst ist deine Adresse bestätigt. Einfach löschen.`;
  if (EINSTELLUNGEN.nurMitAbmeldeKopf && !k["list-unsubscribe"]) return "Kein Abmelde-Kopf, vermutlich Spam. Keine Antwort, einfach löschen.";

  const log = protokollLesen();
  const frueher = log[absender] ? new Date(log[absender]) : null;
  if (frueher && (Date.now() - frueher.getTime()) / 864e5 < EINSTELLUNGEN.erneutNachTagen) {
    return `Diesem Absender hast du schon am ${datum(frueher)} widersprochen.`;
  }

  // Eigene Adresse, an die die Werbung ging (Plus-Adresse aus dem Kopf bevorzugt)
  const kopfEmpf = adressenImText([k["delivered-to"], k["x-original-to"], k["to"]].join(" "));
  const meine = kopfEmpf.find((a) => a.includes("+"))
    || (item.to || []).map((r) => r.emailAddress.toLowerCase()).find((a) => basisDomain(a) === basisDomain(meine0))
    || meine0;

  const text = (await asyncCall((cb) => item.body.getAsync(Office.CoercionType.Text, cb))) + " " +
               (await asyncCall((cb) => item.body.getAsync(Office.CoercionType.Html, cb)));
  const ds = besteAdresseAusMail(text, absender, meine);
  const replyTo = adressenImText(k["reply-to"]).find((a) => !istNoReply(a)) || "";
  let an = replyTo || (!istNoReply(absender) ? absender : "") || ds;
  if (!an && EINSTELLUNGEN.fallbackInfo) an = "info@" + basisDomain(absender);
  if (!an) return "Keine erreichbare Adresse gefunden.";

  const cc = [];
  if (ds && ds !== an) cc.push(ds);
  const abm = EINSTELLUNGEN.abmeldeAdresseInCc ? abmeldeAdresse(k["list-unsubscribe"]) : "";
  if (abm && abm !== an && abm !== ds) cc.push(abm);

  const name = EINSTELLUNGEN.meinName || Office.context.mailbox.userProfile.displayName;
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: [an],
    ccRecipients: cc,
    subject: frueher
      ? "Erneuter Widerspruch und letzte Aufforderung zur Unterlassung von Werbe-E-Mails"
      : "Widerspruch gegen Werbe-E-Mails und Aufforderung zur Unterlassung",
    htmlBody: brieftext(item, meine, frueher, name),
  });

  await protokollSchreiben(absender);
  await kategorieSetzen(item);
  return "";
}

async function widerspruch(event) {
  try {
    const grund = await pruefenUndErstellen();
    if (grund) melden(grund, true);
    else melden("Widerspruch als Entwurf geöffnet. Bitte prüfen und senden.", false);
  } catch (e) {
    melden("Fehler: " + (e && e.message ? e.message : e), true);
  }
  event.completed();
}

window.widerspruch = widerspruch;
Office.onReady(() => {
  if (Office.actions && Office.actions.associate) Office.actions.associate("widerspruch", widerspruch);
});
