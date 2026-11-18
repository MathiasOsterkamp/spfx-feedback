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
Der **Ansprechpartner ist Pflicht** (Empfänger der Mail), das **Funktionspostfach optional** (erhält eine Kopie). Ist nichts
hinterlegt, steht ein Hinweis unter dem Feld.
„Senden“ prüft das Formular und zeigt zuerst eine Vorschau der E-Mail, so wie sie verschickt wird: oben Betreff, An, Cc (immer
sichtbar, ggf. „keine“) und Antworten an, jede Adresse einzeln; darunter der Text. Absender (von SharePoint festgelegt) und Bcc werden nicht angezeigt. Links in der Vorschau öffnen sich in einem neuen Fenster. „Bearbeiten“ führt zurück zum Formular, die Eingaben
bleiben erhalten. Erst „Jetzt senden“ in der Vorschau verschickt genau diese E-Mail über `SP.Utilities.Utility.SendEmail`: **An** den Ansprechpartner, das **Funktionspostfach in Cc**
(zusammen mit `mail.cc` aus der Konfiguration). Die Mail enthält zusätzlich Name und
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

`gulp serve` läuft mit `--locale=de-de` (Task in `.vscode/tasks.json` und `pnpm run serve`): Im Debug-Modus liefert SPFx nur die
Sprachdatei dieser Sprache aus, ohne den Schalter immer `en-us`. Für eine andere Sprache den Schalter ändern (z. B. `--locale=fr-fr`).
Kategorien und Mailvorlage richten sich dagegen nach der Sprache des SharePoint-Benutzers.

Die Debug-URL (`?loadSPFX=true&debugManifestsFile=…&customActions=…`) baut `tools/open-debug-page.js` aus dem Manifest und den
Eigenschaften in `sharepoint/assets/elements.xml` (nur noch `logLevel`), dabei mit `logLevel: verbose`. Beim ersten Aufruf „Debugskripts laden“ bestätigen.

`config/serve.json` enthält bewusst keinen `api`-Server: Er dient nur der lokalen Workbench, und sein Port 5432 ist der
Standardport von PostgreSQL.

## Bereitstellen

1. `spfx-feedback.sppkg` in den **App-Katalog** hochladen und bereitstellen.
2. Auf der Zielwebsite: *Website-Inhalte → App hinzufügen → spfx-feedback-client-side-solution*.
   Das Feature registriert dabei die Custom Action aus `sharepoint/assets/elements.xml`, der Button erscheint auf allen Websiteseiten.
3. Dialog auf einer Seite öffnen → **Zahnrad** → Tab **Einrichtung** → **Einrichten**: legt Einstellungsliste, Liste mit
   Ansprechpartnern und Spalten an (siehe unten). Danach unter **Empfänger** das Funktionspostfach eintragen, bei Bedarf Kategorien und
   Mailvorlage anpassen und speichern. Personen in die Liste eintragen und in den Seiteneigenschaften auswählen.

**E-Mail-Versand:** Auf SharePoint 2019 muss in der Zentraladministration ausgehende E-Mail (SMTP) konfiguriert sein.
Absender ist die dort hinterlegte Adresse. Auf SharePoint Online stellt `SendEmail` nur an Empfänger innerhalb der Organisation zu.

## Konfiguration

Zwei Ebenen:

1. **Code** (`src/extensions/feedback/config/feedbackConfig.ts`, mit dem Paket ausgerollt): alle Einstellungen und die Standardwerte.
2. **Einstellungsliste der Website** („Feedback-Einstellungen“, angelegt und bearbeitet über das Zahnrad im Dialog): überschreibt pro Website
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
| `languages` | Sprachen mit `label` (Name in der Sprachauswahl der Einstellungen), `categories`, `mailSubject` und `mailBody`, siehe unten | `de-de`, `en-us`, `fr-fr` | je Sprache: `Kategorien (de-de)` (eine pro Zeile), `Mail-Betreff (de-de)`, `Mail-Text (de-de)` |
| `mail.cc`, `bcc`, `from`, `replyToReporter` | Weitere Teile der Mail, siehe unten | – | – |
| `defaultLanguage` | Sprache, die gilt, wenn die Oberflächensprache zu keiner Sprache passt | `de-de` | – |

Nicht in die Einstellungsliste gehören: die Einstellungsliste selbst und der Speicherort des Postfachs (die Extension braucht sie, um
die Liste zu finden), die E-Mail-Spalte und die Sicherheitsregeln (`securityConfig.ts`, Websitebesitzer sollen sie nicht lockern können).

**Sprachen.** Kategorien und Mailvorlage werden pro Sprache gepflegt. Die Standardwerte stehen in `languages` in `feedbackConfig.ts`
(Kategorien in der Reihenfolge wie im Formular). Jede Website kann sie pro Sprache überschreiben: Der Titel des Eintrags in der
Einstellungsliste trägt den Sprachcode, z. B. `Kategorien (en-us)`. Fehlt der Eintrag einer Sprache, gilt deren Standardwert aus dem Code.
Eine weitere Sprache kommt mit einem neuen Eintrag in `languages` dazu; für die Texte der Oberfläche braucht sie zusätzlich eine
Sprachdatei in `loc/`.

**Einstellungsliste.** Ein Eintrag pro Einstellung: Titel = Schlüssel, Wert in der Spalte `Wert` (mehrzeiliger Text ohne Formatierung).
Die Liste ist in den Websiteinhalten **ausgeblendet** (`Hidden`); direkt erreichbar bleibt sie über ihre URL
(`<Website>/Lists/FeedbackEinstellungen`). „Einrichten“ legt **alle Einträge** an, vorbelegt mit
den Standardwerten aus dem Code (das Funktionspostfach leer). Spätere Änderungen der Standardwerte im Code wirken deshalb nur noch auf
Einträge, die leer sind.
Gelesen wird sie beim Öffnen des Dialogs mit einem Request und für den Seitenaufruf zwischengespeichert, nicht bei jedem Seitenaufbau.
Gibt es Liste oder Spalte nicht, gelten die Werte aus dem Code. Ungültige Einträge (z. B. ein Spaltenname mit Leerzeichen) werden
ignoriert und in der Konsole gemeldet. Gibt es einen Titel mehrfach, zählt der älteste Eintrag.

Die Sprache (Kategorien und Mailvorlage) wird anhand der Oberflächensprache gewählt: zuerst exakt (`de-de`), dann über das
Sprachpräfix (`de-at` → `de-de`), sonst `defaultLanguage`. Ungültige Spalten- oder Listennamen werden beim Start als Fehler in der Browser-Konsole gemeldet.
Der Ansprechpartner ist dann abgeschaltet, und der Tab „Einrichtung“ zeigt „Konfiguration im Code ungültig“.

**Mailvorlage (`languages`, `mail`).** Diese Teile von `SP.Utilities.Utility.SendEmail` lassen sich festlegen:

| Einstellung | Bedeutung |
|---|---|
| `languages['de-de'].mailSubject` | Betreff (Text) je Sprache. Zeilenumbrüche werden entfernt, höchstens 255 Zeichen. Pro Website und Sprache überschreibbar (`Mail-Betreff (de-de)`). |
| `languages['de-de'].mailBody` | Inhalt als reiner Text ohne HTML, Zeilenumbrüche mit `\n`. Eine Vorlage mit HTML-Tags wird nicht verwendet. Die Hilfsfunktion `mailBody()` erzeugt den Standardtext, ein eigener String geht genauso. Pro Website und Sprache überschreibbar (`Mail-Text (de-de)`, höchstens 20 000 Zeichen). |
| `cc`, `bcc` | Feste zusätzliche Empfänger, z. B. ein Archivpostfach. Sie unterliegen `allowedRecipientDomains`. |
| `from` | Absender. Leer = Absender von SharePoint. SP2019: Der SMTP-Server muss diesen Absender zulassen. SPO: wird in der Regel ignoriert. |
| `replyToReporter` | Header `Reply-To` = meldende Person, „Antworten“ geht also direkt an sie |

Platzhalter in Betreff und Inhalt: `{{category}}`, `{{subject}}`, `{{description}}`, `{{url}}`, `{{reporterName}}`, `{{reporterEmail}}`,
`{{siteTitle}}`, `{{siteUrl}}`, `{{date}}` (in der Oberflächensprache) und `{{recipients}}` (Ansprechpartner und Funktionspostfach). `SendEmail` verschickt HTML, deshalb
wandelt der Code den Inhalt beim Versand um: Vorlagentext und alle Werte werden HTML-escaped, Zeilenumbrüche werden zu `<br>`,
`{{url}}`, `{{siteUrl}}` und `{{reporterEmail}}` werden zu Links, die Schrift entspricht der SharePoint-Oberfläche. Weder Vorlage noch
Eingaben können also HTML einschleusen. Unbekannte Platzhalter werden entfernt und in der Konsole gemeldet. Ohne brauchbare Vorlage gilt
eine minimale Ersatzvorlage.
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
| Ändern im Zahnrad | Ab Entwurfsrechten, sobald die Einstellungsliste eingerichtet ist | Nur Besitzer |
| Technik | REST, auch in SharePoint Online ohne benutzerdefinierte Skripts | CSOM, in SharePoint Online nur mit erlaubten benutzerdefinierten Skripts |
| Direkt ändern | In der Liste, auch durch Mitglieder mit Bearbeiten-Rechten. Wer das nicht will, schränkt die Berechtigungen der Liste ein. | PowerShell, siehe unten |

Bei `list` brauchen Besucher Leserechte auf der Einstellungsliste (bei vererbten Rechten gegeben). Gelesen wird über
`web/getlist(...)`, das gibt es ab SharePoint 2016. Ein Wechsel des Speicherorts übernimmt den bisherigen Wert nicht.

Nur bei `propertyBag`: Funktionspostfach per PowerShell setzen:

```powershell
Connect-PnPOnline -Url https://intranet/sites/team -CurrentCredentials
Set-PnPPropertyBagValue -Key 'Funktionspostfach' -Value 'intranet@firma.de'   # SPO ggf. vorher: Set-PnPNoScriptSite -Disable
```

## Einstellungen im Dialog (Zahnrad)

Personen mit mindestens **Entwerfen**-Rechten auf der Website sehen oben rechts im Dialog ein **Zahnrad**. Geprüft wird die Berechtigung
`ManageLists`, die zur Berechtigungsstufe Entwerfen gehört. Besucher sehen kein Zahnrad, nur den Hinweis unter dem Empfängerfeld.

**Nur hier wird etwas in SharePoint angelegt.** Das Formular selbst zeigt keinen Einrichtungshinweis mehr, und „Speichern“ legt keine
Listen oder Spalten an. Das macht ausschließlich „Einrichten“ im Tab **Einrichtung**.

Die Einstellungen haben drei Tabs. Beim Öffnen ist „Einrichtung“ vorausgewählt, wenn etwas fehlt, sonst „Empfänger“.
Ein Tab mit Fehlern oder fehlender Einrichtung zeigt ein Warnsymbol.

| Tab | Inhalt |
|---|---|
| **Empfänger** | **Ansprechpartner:** Liste mit Ansprechpartnern (Dropdown mit den benutzerdefinierten Listen der Website; Textfeld, falls die Listen nicht gelesen werden können), Nachschlagespalte (Dropdown mit den Nachschlagespalten von „Websiteseiten“: sichtbar, nicht schreibgeschützt, ohne Systemspalten; die Standardspalte aus dem Code steht immer oben, auch wenn es sie noch nicht gibt; Textfeld, falls die Spalten nicht gelesen werden können), Ansprechpartner dieser Seite mit Link „Seiteneigenschaften bearbeiten“. Gibt es die Nachschlagespalte schon und zeigt sie auf eine andere Liste, steht ein Hinweis darunter: Eine andere Liste wirkt erst mit einer neuen Spalte. **Funktionspostfach:** Textfeld mit dem gespeicherten Wert. Bei `list` reichen Entwurfsrechte, bei `propertyBag` nur Besitzer (`ManageWeb`), sonst ist das Feld schreibgeschützt. |
| **Formular und E-Mail** | Sprachauswahl (vorausgewählt ist die Sprache der Oberfläche; angepasste Sprachen sind markiert), darunter **Kategorien** sowie **Betreff** und **Text** der Benachrichtigung in dieser Sprache (reiner Text, HTML-Tags werden abgelehnt). Unter den Feldern: „Standardwert zum Bearbeiten einfügen“ (leeres Feld) bzw. „Standardwert wiederherstellen“ (eigener Wert). Als „angepasst“ gilt eine Sprache, deren Werte vom Standard abweichen. |
| **Einrichtung** | Checkliste: Einstellungsliste (ausgeblendet; fehlen Wertspalte oder Einträge, steht sie auf „unvollständig“), Liste mit Ansprechpartnern, E-Mail-Spalte, Nachschlagespalte in „Websiteseiten“, jeweils *vorhanden*, *fehlt*, *sichtbar*, *unvollständig* oder *Fehler*. **Einrichten** legt alles Fehlende in einem Schritt an, **Erneut prüfen** liest den Stand neu. |

Leere Felder zeigen grau den Standardwert aus dem Code und speichern nichts. **Speichern** unten rechts schreibt alle geänderten Felder
aller Tabs und Sprachen in die Einstellungsliste (bzw. das Postfach in den Property Bag). Vorher geprüft werden Listen- und Spaltennamen
(nur erlaubte Zeichen), Anzahl und Länge der Kategorien, Länge der Vorlage und unbekannte Platzhalter. Bei einem Fehler springt die
Ansicht zu Tab und Sprache des ersten fehlerhaften Felds. Solange die Einstellungsliste nicht eingerichtet ist, sind die Felder
schreibgeschützt; ein Hinweis verlinkt auf den Tab „Einrichtung“. „Zurück zum Formular“ fragt nach, bevor ungespeicherte Änderungen
verworfen werden.

**Einrichten** arbeitet mit den *gespeicherten* Einstellungen. Nach einer Änderung von Liste oder Spalte also erst speichern; bis dahin
ist „Einrichten“ gesperrt. Die Schritte:

1. Einstellungsliste (Vorlage 100, ausgeblendet; eine vorhandene sichtbare Liste wird ausgeblendet) und Wertspalte (mehrzeiliger Text
   ohne Formatierung, in der Standardansicht)
2. Fehlende Einträge der Einstellungsliste, vorbelegt mit den Standardwerten aus dem Code; vorhandene Einträge bleiben unverändert
3. Liste mit Ansprechpartnern, außer die Nachschlagespalte existiert schon, dann wird deren Zielliste verwendet
4. E-Mail-Spalte in dieser Liste
5. Nachschlagespalte in „Websiteseiten“ (allen Inhaltstypen, also auch in den Seiteneigenschaften sichtbar)

Wurde die Liste gelöscht, auf die eine vorhandene Nachschlagespalte zeigt, meldet die Checkliste „verweist auf eine gelöschte Liste“.
„Einrichten“ löscht diese verwaiste Spalte dann und legt sie für die Liste mit Ansprechpartnern neu an (das Ziel einer Nachschlagespalte
lässt sich nicht ändern); die Seiten müssen ihren Ansprechpartner danach neu auswählen. Gelöschte Listen lassen die Prüfung nie
scheitern, sie gelten als fehlend.

Jeder Schritt prüft zuerst, ob es seinen Teil schon gibt. Nach einem Abbruch lässt sich die Einrichtung einfach wiederholen; die Checkliste
zeigt danach den erreichten Stand. Gibt es in „Websiteseiten“ schon eine Spalte mit dem Namen, die keine Nachschlagespalte ist, bleibt sie
unverändert. Die Checkliste meldet das, und unter „Empfänger“ lässt sich ein anderer Spaltenname einstellen. Listen werden über ihre URL
gefunden, nicht über den Titel. Zum Einrichten sind `ManageLists` auf der Website und auf „Websiteseiten“ nötig. Nach jeder Aktion liest
der Dialog die Werte neu.

**Berechtigungen der Einstellungsliste:** Die Liste erbt die Rechte der Website. Damit können auch Mitglieder (Bearbeiten) die Einträge
direkt in der Liste ändern, also auch den Text der Benachrichtigung. Wer das nicht will, unterbricht die Vererbung der Liste und gibt
Mitgliedern nur Leserechte. Besucher brauchen Leserechte.

Technik und Voraussetzungen:
- Liste: REST `web/lists` (Vorlage 100). Die URL entsteht aus dem Titel beim Anlegen, ein abweichender Titel und `Hidden` (nur
  Einstellungsliste) werden danach per MERGE gesetzt.
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
- Längenbegrenzungen (`maxTextLength`, `maxDescriptionLength`): **alle einzeiligen Textfelder höchstens 255 Zeichen** (Betreff, URL,
  Ansprechpartner, Funktionspostfach; in den Einstellungen u. a. Funktionspostfach und Mail-Betreff), **Beschreibung höchstens 5000 Zeichen**.
  Die Felder lassen keine längere Eingabe zu, die Prüfung vor dem Senden meldet zu lange vorbelegte Werte. Dazu: je Adresse 254 Zeichen,
  Anzahl und Länge der Kategorien.
- Mailvorlagen sind reiner Text; Vorlagen mit HTML-Tags werden abgelehnt bzw. ignoriert. Vorlagentext und alle Eingaben werden im Mailtext HTML-escaped. Steuerzeichen und Zeilenumbrüche werden aus dem Betreff entfernt.
  Die URL muss mit `http://` oder `https://` beginnen.

## Projektstruktur

```
├── config/                              Bundle, Paket (Feature mit elements.xml), Serve
├── sharepoint/assets/elements.xml       Custom Action (nur logLevel)
└── src/extensions/feedback/
    ├── FeedbackApplicationCustomizer.ts Einstieg: Konfiguration prüfen, Rechte, Websiteseite erkennen, Button rendern
    ├── components/Feedback.tsx          Schwebender Button
    ├── components/FeedbackDialog.tsx    Formular, Validierung, Vorschau, Senden
    ├── components/MailPreview.tsx       Vorschau der E-Mail vor dem Senden
    ├── components/styles.ts             Styles (Theme der Website)
    ├── config/feedbackConfig.ts         KONFIGURATION: Spalten, Listen, Schlüssel der Einstellungsliste, Mailvorlage
    ├── config/securityConfig.ts         Zentrale Sicherheitsregeln
    ├── services/MailService.ts          SharePoint-REST SendEmail (To, CC, BCC, From, Reply-To)
    ├── services/SettingsService.ts      Einstellungsliste lesen und auf die Code-Konfiguration anwenden
    ├── services/RecipientService.ts     Ansprechpartner (Seite) und Funktionspostfach (Liste oder Property Bag), mit Status
    ├── services/SetupService.ts         Einrichtung prüfen und anlegen, Einstellungen speichern, Postfach im Property Bag (CSOM)
    ├── components/SettingsView.tsx      Einstellungen und Einrichtung (Zahnrad, ab Entwurfsrechten)
    ├── utils/mail.ts                    Mail aus der Vorlage (Platzhalter, Escaping)
    ├── utils/validation.ts              Pflichtfelder, URL, E-Mail
    ├── utils/setupValidation.ts         Prüfung der Felder der Einstellungen
    ├── utils/sanitize.ts                Prüfung der Konfiguration und der Einstellungsliste
    ├── utils/logging.ts / theme.ts / untyped.ts   aus der Standardlösung
    └── loc/                             Texte der Oberfläche: de-de, en-us, fr-fr
```

## Barrierefreiheit

- Button mit `aria-label`/Tooltip, sichtbarem Fokus und Hochkontrast-Darstellung, wird beim Drucken ausgeblendet.
- Dialog mit Fokusfalle; Spalten als `<fieldset>`/`<legend>`; Pflichtfelder mit `required`.
- Fehlermeldungen erst nach dem ersten Absenden, danach live. Der Fokus springt auf das erste fehlerhafte Feld.
- Klick neben den Dialog verwirft keine Eingaben (`isBlocking`). Enter in einem einzeiligen Feld öffnet die Vorschau; dort liegt der Fokus auf „Jetzt senden“.
