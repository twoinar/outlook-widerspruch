/* Werbe-Widerspruch: Outlook-Add-in (neues Outlook, Outlook im Web, Mac)
 *
 * Schaltfläche "Widerspruch" in einer geöffneten Werbemail:
 *  - prüft, ob es Spam ist (dann keine Antwort, sonst wird die Adresse bestätigt)
 *  - sucht eine erreichbare Adresse (Reply-To, Absender, Datenschutz/Impressum, info@)
 *  - sendet den Widerspruch, verschiebt alle Mails des Absenders nach "Werbung"
 *    und legt eine Regel an, die künftige Mails des Absenders gleich dorthin schiebt
 *  - ohne Microsoft-Graph-Anmeldung (clientId leer): öffnet nur den Entwurf
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

  // --- Microsoft Graph (App-Anmeldung, siehe Anleitung). Leer = nur Entwurf öffnen ---
  clientId: "33b78f44-4e96-4741-870b-838fd42e728c",          // Anwendungs-ID (Client) der App-Anmeldung
  // "common" = jedes Postfach: private Microsoft-Konten und alle Firmen-Umgebungen
  authority: "https://login.microsoftonline.com/common",
  sofortSenden: true,    // false = Entwurf öffnen statt senden, Aufräumen läuft trotzdem
  aufraeumen: true,      // alle Mails des Absenders aus dem Posteingang verschieben
  regelAnlegen: true,    // Regel: künftige Mails des Absenders automatisch verschieben
  zielOrdner: "Werbung",
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

// ---------- Sprache ----------
const WOERTER_DE = ["und", "der", "die", "das", "sie", "ihre", "ihr", "mit", "für", "nicht", "auf", "ist", "jetzt", "hier", "abmelden", "newsletter-abmeldung"];
const WOERTER_EN = ["the", "and", "your", "you", "with", "for", "not", "this", "now", "here", "unsubscribe", "our", "is", "are"];

// "de" oder "en": Content-Language-Kopf, sonst Wortzählung im Text, bei Gleichstand die Absenderdomain
function sprache(k, text, absender) {
  const cl = (k["content-language"] || "").toLowerCase();
  if (cl.startsWith("de")) return "de";
  if (cl.startsWith("en")) return "en";
  const woerter = (text || "").replace(/<[^>]+>/g, " ").toLowerCase().match(/[a-zäöüß-]+/g) || [];
  let de = 0, en = 0;
  for (const w of woerter.slice(0, 3000)) {
    if (WOERTER_DE.includes(w)) de++;
    if (WOERTER_EN.includes(w)) en++;
  }
  if (de > en * 1.2) return "de";
  if (en > de * 1.2) return "en";
  return /\.(de|at|ch)$/.test(absender) ? "de" : "en";
}

function betreffText(lang, frueher) {
  if (lang === "en") {
    return frueher
      ? "Renewed objection and final demand to cease sending advertising e-mails"
      : "Objection to advertising e-mails and demand to cease and desist";
  }
  return frueher
    ? "Erneuter Widerspruch und letzte Aufforderung zur Unterlassung von Werbe-E-Mails"
    : "Widerspruch gegen Werbe-E-Mails und Aufforderung zur Unterlassung";
}

function brieftextEn(item, meineAdr, frueher, name) {
  const d = (x) => x.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" });
  const absaetze = [];
  absaetze.push(esc("Dear Sir or Madam,"));
  absaetze.push(esc(`On ${d(item.dateTimeCreated)} you sent an advertising e-mail with the subject "${item.subject}" to my e-mail address ${meineAdr} (sender: ${item.from.emailAddress}).`));
  if (frueher) {
    absaetze.push(esc(`By e-mail of ${d(frueher)} I already informed you that I do not wish to receive any advertising from you and demanded that you stop. Sending me advertising again nevertheless is a deliberate violation that I will not tolerate any further.`));
  }
  absaetze.push(esc("I never subscribed to your newsletter and have at no time given my express consent to receiving advertising by e-mail. Sending it is unlawful under Art. 13 of the ePrivacy Directive 2002/58/EC and, under German law, constitutes an unreasonable nuisance within the meaning of Section 7(2) No. 2 of the German Act against Unfair Competition (UWG) and an unlawful interference with my general right of personality (Sections 823(1), 1004(1) German Civil Code). Insofar as you rely on an existing customer relationship, I hereby expressly object to any further use of my address."));
  const punkte = [
    "to refrain immediately from sending any advertising to me, regardless of which of my e-mail addresses;",
    "to erase my personal data (Art. 17 GDPR). I object to the processing of my data for direct marketing purposes under Art. 21(2) GDPR; it may no longer be processed for such purposes (Art. 21(3) GDPR);",
  ];
  if (EINSTELLUNGEN.mitAuskunft) {
    punkte.push("to inform me under Art. 15 GDPR which personal data you hold about me, where you obtained my address and to whom you have disclosed it. If you rely on consent, please provide proof of it (Art. 7(1) GDPR);");
  }
  punkte.push("to confirm in writing within 14 days that you have complied.");
  absaetze.push("I therefore request you<br>" + punkte.map((p, i) => `${i + 1}. ${esc(p)}`).join("<br>"));
  if (frueher) {
    absaetze.push(esc("Should I not receive confirmation within this period, or should I receive further advertising from you, I will instruct a lawyer to issue a formal warning (Abmahnung) without further notice and file a complaint with the competent data protection supervisory authority. You will bear the costs."));
  } else {
    absaetze.push(esc("In the event of any further unsolicited e-mails I expressly reserve the right to take legal action, in particular a formal warning with costs and a complaint to the competent data protection supervisory authority."));
  }
  absaetze.push("Yours faithfully,<br>" + esc(name));
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

// ---------- Microsoft Graph ----------
const SCOPES = ["Mail.ReadWrite", "Mail.Send", "MailboxSettings.ReadWrite"];
let msalApp = null;

function graphAktiv() {
  return !!(EINSTELLUNGEN.clientId && window.msal && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
}

async function token() {
  if (!msalApp) {
    msalApp = await msal.createNestablePublicClientApplication({
      auth: {
        clientId: EINSTELLUNGEN.clientId,
        authority: EINSTELLUNGEN.authority,
      },
    });
  }
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (e) {
    return (await msalApp.acquireTokenPopup({ scopes: SCOPES })).accessToken;
  }
}

async function graph(methode, pfad, daten) {
  const url = pfad.startsWith("https://") ? pfad : "https://graph.microsoft.com/v1.0" + pfad;
  const r = await fetch(url, {
    method: methode,
    headers: { Authorization: "Bearer " + (await token()), "Content-Type": "application/json" },
    body: daten ? JSON.stringify(daten) : undefined,
  });
  if (!r.ok) throw new Error(`Graph ${r.status}: ${(await r.text()).slice(0, 120)}`);
  return r.status === 202 || r.status === 204 ? null : r.json();
}

function odataText(s) {
  return encodeURIComponent(s.replace(/'/g, "''"));
}

async function zielOrdnerId() {
  const name = odataText(EINSTELLUNGEN.zielOrdner);
  for (const pfad of [`/me/mailFolders/inbox/childFolders?$filter=displayName eq '${name}'`,
                      `/me/mailFolders?$filter=displayName eq '${name}'`]) {
    const r = await graph("GET", pfad);
    if (r.value.length) return r.value[0].id;
  }
  return (await graph("POST", "/me/mailFolders/inbox/childFolders", { displayName: EINSTELLUNGEN.zielOrdner })).id;
}

async function regelAnlegen(absender, ordnerId) {
  const titel = "Widerspruch: " + absender;
  const r = await graph("GET", "/me/mailFolders/inbox/messageRules");
  if (r.value.some((x) => x.displayName === titel)) return false;
  await graph("POST", "/me/mailFolders/inbox/messageRules", {
    displayName: titel,
    sequence: 1,
    isEnabled: true,
    conditions: { senderContains: [absender] },
    actions: { moveToFolder: ordnerId, stopProcessingRules: true },
  });
  return true;
}

async function posteingangAufraeumen(absender, ordnerId, ausser) {
  let n = 0;
  let url = `/me/mailFolders/inbox/messages?$filter=from/emailAddress/address eq '${odataText(absender)}'&$select=id&$top=100`;
  while (url && n < 1000) {
    const r = await graph("GET", url);
    for (const m of r.value) {
      if (m.id === ausser) continue;
      await graph("POST", `/me/messages/${m.id}/move`, { destinationId: ordnerId });
      n++;
    }
    url = r["@odata.nextLink"];
  }
  return n;
}

// ---------- Hauptablauf ----------
// Ergebnis: { fehler } bei Abbruch, sonst { text } mit Zusammenfassung
async function ausfuehren() {
  const item = Office.context.mailbox.item;
  const meine0 = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  const absender = (item.from && item.from.emailAddress || "").toLowerCase();
  if (!absender.includes("@")) return { fehler: "Keine Absenderadresse gefunden." };
  if (basisDomain(absender) === basisDomain(meine0)) return { fehler: "Interner Absender, kein Widerspruch." };

  const k = kopfZerlegen(await asyncCall((cb) => item.getAllInternetHeadersAsync(cb)));
  const scl = spamWert(k);
  if (scl >= EINSTELLUNGEN.spamAbScl) return { fehler: `Sieht nach Spam aus (SCL ${scl}). Keine Antwort, sonst ist deine Adresse bestätigt. Einfach löschen.` };
  if (EINSTELLUNGEN.nurMitAbmeldeKopf && !k["list-unsubscribe"]) return { fehler: "Kein Abmelde-Kopf, vermutlich Spam. Keine Antwort, einfach löschen." };

  let mitGraph = graphAktiv();
  const log = protokollLesen();
  const frueher = log[absender] ? new Date(log[absender]) : null;
  const gesperrt = frueher && (Date.now() - frueher.getTime()) / 864e5 < EINSTELLUNGEN.erneutNachTagen;
  if (gesperrt && !mitGraph) return { fehler: `Diesem Absender hast du schon am ${datum(frueher)} widersprochen.` };

  const teile = [];
  if (!gesperrt) {
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
    if (!an) return { fehler: "Keine erreichbare Adresse gefunden." };

    const cc = [];
    if (ds && ds !== an) cc.push(ds);
    const abm = EINSTELLUNGEN.abmeldeAdresseInCc ? abmeldeAdresse(k["list-unsubscribe"]) : "";
    if (abm && abm !== an && abm !== ds) cc.push(abm);

    const name = EINSTELLUNGEN.meinName || Office.context.mailbox.userProfile.displayName;
    const lang = sprache(k, text, absender);
    const betreff = betreffText(lang, frueher);
    const html = lang === "en" ? brieftextEn(item, meine, frueher, name) : brieftext(item, meine, frueher, name);

    let gesendet = false;
    if (mitGraph && EINSTELLUNGEN.sofortSenden) {
      try {
        await graph("POST", "/me/sendMail", {
        message: {
          subject: betreff,
          body: { contentType: "HTML", content: html },
          toRecipients: [an].map((a) => ({ emailAddress: { address: a } })),
          ccRecipients: cc.map((a) => ({ emailAddress: { address: a } })),
        },
        saveToSentItems: true,
        });
        gesendet = true;
        teile.push(`Widerspruch (${lang === "en" ? "englisch" : "deutsch"}) an ${an} gesendet`);
      } catch (e) {
        // z.B. Firmenpostfach ohne Zustimmung des Admins: dann wie bisher als Entwurf
        console.log("Graph nicht verfügbar:", e);
        mitGraph = false;
      }
    }
    if (!gesendet) {
      Office.context.mailbox.displayNewMessageForm({ toRecipients: [an], ccRecipients: cc, subject: betreff, htmlBody: html });
      teile.push(`Widerspruch (${lang === "en" ? "englisch" : "deutsch"}) als Entwurf geöffnet`);
    }
    await protokollSchreiben(absender);
    await kategorieSetzen(item);
  } else {
    teile.push(`Widerspruch schon am ${datum(frueher)} gesendet`);
  }

  if (!mitGraph) return { text: teile[0] + ". Bitte prüfen und senden." };

  // Aufräumen: Regel für künftige Mails, Posteingang leeren, diese Mail zuletzt verschieben
  let ordnerId;
  try {
    ordnerId = await zielOrdnerId();
  } catch (e) {
    return { text: teile[0] + ". Aufräumen ist für dieses Postfach nicht freigegeben." };
  }
  if (EINSTELLUNGEN.regelAnlegen && await regelAnlegen(absender, ordnerId)) teile.push("Regel angelegt");
  if (EINSTELLUNGEN.aufraeumen) {
    const diese = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    const n = await posteingangAufraeumen(absender, ordnerId, diese);
    teile.push(`${n + 1} Mail(s) nach ${EINSTELLUNGEN.zielOrdner} verschoben`);
    melden(teile.join(", ") + ".", false);
    await graph("POST", `/me/messages/${diese}/move`, { destinationId: ordnerId });
  }
  return { text: teile.join(", ") + "." };
}

async function widerspruch(event) {
  try {
    const r = await ausfuehren();
    if (r.fehler) melden(r.fehler, true);
    else melden(r.text, false);
  } catch (e) {
    melden("Fehler: " + (e && e.message ? e.message : e), true);
  }
  event.completed();
}

window.widerspruch = widerspruch;
Office.onReady(() => {
  if (Office.actions && Office.actions.associate) Office.actions.associate("widerspruch", widerspruch);
});
