# SPFx-Lösung „Feedback geben / Inhalte melden“

SharePoint-Framework-Erweiterung (Application Customizer, SPFx **1.4.1**, passend für SharePoint 2019) mit **React 16.14**,
**Fluent UI 8** und **TypeScript 4.9**. Basiert auf der Standardlösung `spfx-image-map-solution` (Build, Lint, Logging, Theme).

Auf allen **Seiten der Bibliothek „Websiteseiten“** (Seiten und News-Beiträge) erscheint unten rechts ein runder **Feedback-Button**.
Auf Listen, Bibliotheken, Websiteinhalten usw. erscheint er nicht. Bei Navigation ohne Neuladen wird das neu geprüft. Der Button öffnet
den Dialog „Feedback geben / Inhalte melden“:

| Ihr Anliegen | Inhaltliche Verantwortung |
|---|---|
| Kategorie * (Auswahl) | Ansprechpartner (E-Mail-Adresse): aus der **Nachschlagespalte der Seite** (Liste mit Titel und E-Mail) |
| Betreff * | Funktionspostfach (E-Mail-Adresse): aus der **Einstellungsliste der Website** (oder dem Property Bag, je nach Konfiguration) || URL * (vorbelegt mit der aktuellen Seiten-URL) | |
| Beschreibung * | |

Die Adressen werden beim Öffnen des Dialogs geladen und vorbelegt. Sie bleiben änderbar, mehrere Adressen werden mit `;` getrennt.
Mindestens eine Adresse ist Pflicht. Ist nichts hinterlegt, steht ein Hinweis unter dem Feld.
„Senden“ verschickt eine E-Mail über `SP.Utilities.Utility.SendEmail` an alle Adressen. Die Mail enthält zusätzlich Name und
E-Mail der meldenden Person, die Website und das Datum.

Die URL ist die Adresse der aktuellen Seite ohne Query-String und Anker. So landen z. B. die Debug-Parameter nicht in der Meldung.

## Voraussetzungen

| Werkzeug | Version |
|---|---|
| Node.js | ≥ 22 (getestet mit 26) |
| pnpm | 10 (`.npmrc`: `shamefully-hoist`, `ignore-scripts`, `--openssl-legacy-provider`) |

```bash
pnpm install
pnpm run typecheck && pnpm run lint
pnpm run ship        # gulp clean && gulp bundle --ship && gulp package-solution --ship
```

Ergebnis: `sharepoint/solution/spfx-feedback.sppkg` (Skripte im Paket enthalten, kein CDN nötig).

## Debuggen

Erweiterungen laufen nicht in der Workbench. Voraussetzung: `.vscode/settings.json` → `spfx.siteUrl` (Vorlage: `settings.example.json`).

| Startkonfiguration (F5) | Was passiert |
|---|---|
| **SharePoint home page (own browser)** | Startet `gulp serve` und öffnet die **Startseite** der Website im **Standardbrowser** (eigenes Profil, bereits angemeldet). Debuggen im Browser: F12 → Quellen → `webpack://` → `src`. |
| SharePoint home page (VS Code debugger, separate Chrome) | Haltepunkte direkt in VS Code. Startet ein eigenes Chrome mit leerem Profil, dort einmal anmelden. |

Die Debug-URL (`?loadSPFX=true&debugManifestsFile=…&customActions=…`) baut `tools/open-debug-page.js` aus dem Manifest und den
Eigenschaften in `sharepoint/assets/elements.xml` (nur noch `logLevel`), dabei mit `logLevel: verbose`. Beim ersten Aufruf „Debugskripts laden“ bestätigen.

`config/serve.json` enthält bewusst keinen `api`-Server: Er dient nur der lokalen Workbench, und sein Port 5432 ist der
Standardport von PostgreSQL.

## Bereitstellen

1. `spfx-feedback.sppkg` in den **App-Katalog** hochladen und bereitstellen.
2. Auf der Zielwebsite: *Website-Inhalte → App hinzufügen → spfx-feedback-client-side-solution*.
   Das Feature registriert dabei die Custom Action aus `sharepoint/assets/elements.xml`, der Button erscheint auf allen Websiteseiten.
3. Dialog auf einer Seite öffnen → **Zahnrad** → bei Bedarf Liste/Spalte, Kategorien und Mailvorlage anpassen, Funktionspostfach
   eintragen und speichern; „Einrichten“ legt Liste und Spalten an (siehe unten). Danach Personen in die Liste eintragen und in den
   Seiteneigenschaften auswählen.

**E-Mail-Versand:** Auf SharePoint 2019 muss in der Zentraladministration ausgehende E-Mail (SMTP) konfiguriert sein.
Absender ist die dort hinterlegte Adresse. Auf SharePoint Online stellt `SendEmail` nur an Empfänger innerhalb der Organisation zu.

## Konfiguration

Zwei Ebenen:

1. **Code** (`src/extensions/feedback/config/feedbackConfig.ts`, mit dem Paket ausgerollt): alle Einstellungen und die Standardwerte.
2. **Einstellungsliste der Website** („Feedback-Einstellungen“, bearbeitet über das Zahnrad im Dialog): überschreibt pro Website
   die wichtigsten Werte. Ein leerer oder fehlender Eintrag bedeutet: Es gilt der Wert aus dem Code.

Die Eigenschaften der Custom Action enthalten nur `logLevel`.

| Einstellung im Code | Bedeutung | Standard | In der Einstellungsliste (Titel) |
|---|---|---|---|
| `contact.pageColumn` | Nachschlagespalte in „Websiteseiten“: `internalName` (nur Buchstaben, Ziffern, `_`) und `displayName` | `Ansprechpartner` | `Ansprechpartner-Spalte` (interner Name) |
| `contact.list` | Liste mit den Ansprechpartnern: `url` (immer `Lists/<Name>`) und `title` | `Lists/Ansprechpartner`, „Ansprechpartner“ | `Ansprechpartner-Liste` (Name in der URL) |
| `contact.emailColumn` | E-Mail-Spalte der Liste (einzeiliger Text): `internalName`, `displayName` | `Email`, „E-Mail“ | – |
| `settingsList` | Einstellungsliste: `url` (immer `Lists/<Name>`), `title`, Wertspalte `valueColumn` und die Titel der Einträge (`keys`) | `Lists/FeedbackEinstellungen`, Spalte `Wert` | – |
| `mailbox.storage` | **Speicherort des Funktionspostfachs:** `list` = Einstellungsliste, `propertyBag` = Property Bag der Website | `list` | – |
| `mailbox.key` | Titel des Eintrags in der Einstellungsliste bzw. Schlüssel im Property Bag | `Funktionspostfach` | `Funktionspostfach` (bei `list`) |
| Sprachdateien `categories` | Kategorien je Sprache | siehe `loc/*.js` | `Kategorien` (eine pro Zeile, für alle Sprachen) |
| `mail.templates` | Betreff und Text je Sprache, siehe unten | Tabelle mit allen Werten | `Mail-Betreff`, `Mail-Text` (für alle Sprachen) |
| `mail.cc`, `bcc`, `from`, `replyToReporter` | Weitere Teile der Mail, siehe unten | – | – |
| `defaultLanguage` | Sprache, deren Mailvorlage gilt, wenn die Oberflächensprache keine hat | `de-de` | – |

Nicht in die Einstellungsliste gehören: die Einstellungsliste selbst und der Speicherort des Postfachs (die Extension braucht sie, um
die Liste zu finden), die E-Mail-Spalte und die Sicherheitsregeln (`securityConfig.ts`, Websitebesitzer sollen sie nicht lockern können).

**Kategorien** stehen standardmäßig in den Sprachdateien `loc/de-de.js`, `loc/en-us.js` und `loc/fr-fr.js` (Schlüssel `categories`,
Reihenfolge wie im Formular). SPFx wählt die Datei nach der Oberflächensprache. Für andere Sprachen nimmt SPFx `en-us`.
Sind in der Einstellungsliste Kategorien eingetragen, gelten sie für alle Sprachen.

**Einstellungsliste.** Ein Eintrag pro Einstellung: Titel = Schlüssel, Wert in der Spalte `Wert` (mehrzeiliger Text ohne Formatierung).
Gelesen wird sie beim Öffnen des Dialogs mit einem Request und für den Seitenaufruf zwischengespeichert, nicht bei jedem Seitenaufbau.
Gibt es Liste oder Spalte nicht, gelten die Werte aus dem Code. Ungültige Einträge (z. B. ein Spaltenname mit Leerzeichen) werden
ignoriert und in der Konsole gemeldet. Gibt es einen Titel mehrfach, zählt der älteste Eintrag.

Die Mailvorlage wird anhand der Oberflächensprache gewählt: zuerst exakt (`de-de`), dann über das Sprachpräfix (`de-at` → `de-de`),
sonst `defaultLanguage`. Ungültige Spalten- oder Listennamen werden beim Start als Fehler in der Browser-Konsole gemeldet.
Der Ansprechpartner ist dann abgeschaltet, und die Einrichtung zeigt „Konfiguration ungültig“.

**Mailvorlage (`mail`).** Diese Teile von `SP.Utilities.Utility.SendEmail` lassen sich festlegen:

| Einstellung | Bedeutung |
|---|---|
| `templates['de-de'].subject` | Betreff (Text) je Oberflächensprache. Zeilenumbrüche werden entfernt, höchstens 255 Zeichen. Pro Website überschreibbar (`Mail-Betreff`). |
| `templates['de-de'].body` | Inhalt als HTML. Nur Inline-Styles verwenden, weil viele Mailprogramme `<style>` ignorieren. Die Hilfsfunktion `mailBody()` erzeugt die Standardtabelle, ein eigener HTML-String geht genauso. Pro Website überschreibbar (`Mail-Text`, höchstens 20 000 Zeichen). |
| `cc`, `bcc` | Feste zusätzliche Empfänger, z. B. ein Archivpostfach. Sie unterliegen `allowedRecipientDomains`. |
| `from` | Absender. Leer = Absender von SharePoint. SP2019: Der SMTP-Server muss diesen Absender zulassen. SPO: wird in der Regel ignoriert. |
| `replyToReporter` | Header `Reply-To` = meldende Person, „Antworten“ geht also direkt an sie |

Platzhalter in Betreff und Inhalt: `{{category}}`, `{{subject}}`, `{{description}}`, `{{url}}`, `{{reporterName}}`, `{{reporterEmail}}`,
`{{siteTitle}}`, `{{siteUrl}}`, `{{date}}` (in der Oberflächensprache) und `{{recipients}}`. Im Inhalt werden alle Werte HTML-escaped,
Zeilenumbrüche der Beschreibung werden zu `<br>`. Eingaben können also kein HTML einschleusen. Unbekannte Platzhalter werden entfernt
und in der Konsole gemeldet. Ohne brauchbare Vorlage gilt eine minimale Ersatzvorlage.
Nicht möglich über diese API: Anhänge, Priorität, Text-Alternative zum HTML.

**Ansprechpartner:** Die Seite verweist über die Nachschlagespalte auf einen oder mehrere Einträge der Liste. Angezeigt wird der Titel,
gelesen wird die E-Mail-Spalte. Eine Zelle darf mehrere Adressen enthalten, getrennt mit `;`. Beim Öffnen des Dialogs laufen
zuerst 1 Request für die Einstellungsliste (enthält bei `list` auch das Funktionspostfach), dann 4 Requests für den Ansprechpartner:
Schema der Spalte, dann parallel Wert und Schema der E-Mail-Spalte, dann die Einträge. Bei `propertyBag` kommt 1 Request für das
Funktionspostfach dazu. Alles wird pro Seitenaufruf zwischengespeichert.

**Funktionspostfach: zwei Speicherorte, fest umgeschaltet über `mailbox.storage`.**

| | `list` (Standard) | `propertyBag` |
|---|---|---|
| Wo | Eintrag mit Titel `mailbox.key` in der Einstellungsliste, Wert in der Spalte `settingsList.valueColumn` | Schlüssel `mailbox.key` im Property Bag der Website |
| Ändern im Zahnrad | Ab Entwurfsrechten. Liste und Spalte werden beim ersten Speichern angelegt. | Nur Besitzer |
| Technik | REST, auch in SharePoint Online ohne benutzerdefinierte Skripts | CSOM, in SharePoint Online nur mit erlaubten benutzerdefinierten Skripts |
| Direkt ändern | In der Liste, auch durch Mitglieder mit Bearbeiten-Rechten. Wer das nicht will, schränkt die Berechtigungen der Liste ein. | PowerShell, siehe unten |

Bei `list` brauchen Besucher Leserechte auf der Einstellungsliste (bei vererbten Rechten gegeben). Gelesen wird über
`web/getlist(...)`, das gibt es ab SharePoint 2016. Ein Wechsel des Speicherorts übernimmt den bisherigen Wert nicht.

Nur bei `propertyBag`: Funktionspostfach per PowerShell setzen:

```powershell
Connect-PnPOnline -Url https://intranet/sites/team -CurrentCredentials
Set-PnPPropertyBagValue -Key 'Funktionspostfach' -Value 'intranet@firma.de'   # SPO ggf. vorher: Set-PnPNoScriptSite -Disable
```

## Einrichtung im Dialog (Zahnrad)

Personen mit mindestens **Entwerfen**-Rechten auf der Website sehen oben rechts im Dialog ein **Zahnrad**. Geprüft wird die Berechtigung
`ManageLists`, die zur Berechtigungsstufe Entwerfen gehört. Fehlt etwas, zeigt das Formular zusätzlich den Hinweis
„Die Einrichtung für diese Website ist unvollständig.“ mit dem Link „Einrichtung“. Besucher sehen weder Zahnrad noch Hinweis,
nur den Hinweis unter dem Feld.

Die Einrichtung ist ein Formular in zwei Spalten:

| Ansprechpartner | Funktionspostfach |
|---|---|
| Felder **Liste mit Ansprechpartnern** (Name in der URL) und **Nachschlagespalte** (interner Name) | Textfeld mit dem gespeicherten Wert |
| Status der Nachschlagespalte, Link zur Liste, Ansprechpartner dieser Seite (Link „Seiteneigenschaften bearbeiten“) | Bei `list` reichen Entwurfsrechte, bei `propertyBag` nur Besitzer (`ManageWeb`), sonst ist das Feld schreibgeschützt. |
| **Einrichten** (nur wenn etwas fehlt): legt die Liste mit „Titel“ und E-Mail-Spalte an (falls nötig), ergänzt eine fehlende E-Mail-Spalte und fügt die Nachschlagespalte in „Websiteseiten“ hinzu (allen Inhaltstypen, also auch in den Seiteneigenschaften sichtbar) | |

| Formular | Benachrichtigung |
|---|---|
| **Kategorien**, eine pro Zeile | **Betreff** und **Text (HTML)** der E-Mail; der Link „Standardvorlage zum Bearbeiten einfügen“ übernimmt die Vorlage der aktuellen Sprache als Ausgangspunkt |

Leere Felder zeigen grau den Standardwert aus dem Code und speichern nichts. **Speichern** unten rechts schreibt alle geänderten Felder
in die Einstellungsliste (bzw. das Postfach in den Property Bag). Liste und Wertspalte werden beim ersten Speichern angelegt. Geprüft wird
vorher: Listen- und Spaltennamen (nur erlaubte Zeichen), Anzahl und Länge der Kategorien, Länge der Vorlage und unbekannte Platzhalter.
Nach einer Änderung von Liste oder Spalte erst speichern, dann „Einrichten“, falls sie noch nicht existieren.

Jeder Schritt von „Einrichten“ prüft zuerst, ob es seinen Teil schon gibt. Nach einem Abbruch lässt sich die Einrichtung einfach wiederholen.
Die Liste wird über ihre URL gefunden, nicht über den Titel. Nach jeder Aktion liest der Dialog die Werte neu.

**Berechtigungen der Einstellungsliste:** Die Liste erbt die Rechte der Website. Damit können auch Mitglieder (Bearbeiten) die Einträge
direkt in der Liste ändern, also auch den HTML-Text der Benachrichtigung. Wer das nicht will, unterbricht die Vererbung der Liste und gibt
Mitgliedern nur Leserechte. Besucher brauchen Leserechte.

Technik und Voraussetzungen:
- Liste: REST `web/lists` (Vorlage 100). Die URL entsteht aus dem Titel beim Anlegen, ein abweichender Titel wird danach gesetzt.
- Spalten: REST `fields/createfieldasxml`. E-Mail: Text, Pflichtfeld, in Standardinhaltstyp und Standardansicht.
  Nachschlagespalte: `ShowField="Title"`, alle Inhaltstypen.
- Funktionspostfach: REST kann den Property Bag nicht schreiben, deshalb ein CSOM-Request an `/_vti_bin/client.svc/ProcessQuery`
  (`web.AllProperties[key] = value; web.Update()`). In **SharePoint Online** schlägt das auf Websites ohne benutzerdefinierte Skripts
  (Standard bei modernen Team- und Kommunikationswebsites) mit „Zugriff verweigert“ fehl. Dann den Wert per PowerShell setzen.
- Was angezeigt wird, entscheidet der Browser anhand der Berechtigungen im Seitenkontext. SharePoint prüft sie beim Speichern erneut.

## Dialog

Responsiv: 90 % der Fensterbreite, höchstens 1100 px; unter 480 px Vollbild. Die Breite wird über die `styles` des Dialogs gesetzt,
nicht über `modalProps.styles`, weil Fluent ab 480 px selbst `width: auto` setzt.

## Sicherheit

`src/extensions/feedback/config/securityConfig.ts` enthält die Sicherheitsregeln:

- **`allowedRecipientDomains`**: Die Empfängerfelder sind für alle Besucher änderbar. Ohne Einschränkung lässt sich das Formular nutzen,
  um Mails mit SharePoint als Absender an beliebige Adressen zu schicken. **Für den Produktivbetrieb die eigene(n) Domain(s) eintragen**,
  z. B. `['firma.de']` (Subdomains eingeschlossen).
- Längenbegrenzungen für Betreff (255), Beschreibung (4000), URL (2048), Adressen (254) und Kategorien.
- Alle Eingaben werden im Mailtext HTML-escaped. Steuerzeichen und Zeilenumbrüche werden aus dem Betreff entfernt.
  Die URL muss mit `http://` oder `https://` beginnen.

## Projektstruktur

```
├── config/                              Bundle, Paket (Feature mit elements.xml), Serve
├── sharepoint/assets/elements.xml       Custom Action (nur logLevel)
└── src/extensions/feedback/
    ├── FeedbackApplicationCustomizer.ts Einstieg: Konfiguration prüfen, Rechte, Websiteseite erkennen, Button rendern
    ├── components/Feedback.tsx          Schwebender Button
    ├── components/FeedbackDialog.tsx    Formular, Validierung, Senden
    ├── components/styles.ts             Styles (Theme der Website)
    ├── config/feedbackConfig.ts         KONFIGURATION: Spalten, Listen, Schlüssel der Einstellungsliste, Mailvorlage
    ├── config/securityConfig.ts         Zentrale Sicherheitsregeln
    ├── services/MailService.ts          SharePoint-REST SendEmail (To, CC, BCC, From, Reply-To)
    ├── services/SettingsService.ts      Einstellungsliste lesen und auf die Code-Konfiguration anwenden
    ├── services/RecipientService.ts     Ansprechpartner (Seite) und Funktionspostfach (Liste oder Property Bag), mit Status
    ├── services/SetupService.ts         Listen und Spalten anlegen, Einstellungen speichern, Postfach im Property Bag (CSOM)
    ├── components/SetupView.tsx         Einrichtung (Zahnrad, ab Entwurfsrechten)
    ├── utils/mail.ts                    Mail aus der Vorlage (Platzhalter, Escaping)
    ├── utils/validation.ts              Pflichtfelder, URL, E-Mail
    ├── utils/setupValidation.ts         Prüfung der Felder der Einrichtung
    ├── utils/sanitize.ts                Prüfung der Konfiguration und der Einstellungsliste
    ├── utils/logging.ts / theme.ts / untyped.ts   aus der Standardlösung
    └── loc/                             Texte und Kategorien: de-de, en-us, fr-fr
```

## Barrierefreiheit

- Button mit `aria-label`/Tooltip, sichtbarem Fokus und Hochkontrast-Darstellung, wird beim Drucken ausgeblendet.
- Dialog mit Fokusfalle; Spalten als `<fieldset>`/`<legend>`; Pflichtfelder mit `required`.
- Fehlermeldungen erst nach dem ersten Absenden, danach live. Der Fokus springt auf das erste fehlerhafte Feld.
- Klick neben den Dialog verwirft keine Eingaben (`isBlocking`). Enter in einem einzeiligen Feld sendet.
