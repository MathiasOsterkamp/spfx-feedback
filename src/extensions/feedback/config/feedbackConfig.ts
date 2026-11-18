/**
 * CONFIGURATION OF THE FEEDBACK BUTTON
 *
 * Rolled out with the package. The values are checked on start; invalid names are reported in the
 * browser console.
 *
 * Some settings can be overridden per site in the settings list (see settingsList). They are edited in
 * the settings view of the dialog (gear icon): contact list and lookup column and - per language - categories,
 * mail subject and mail text. An empty or missing entry there means: the value from this file applies.
 *
 * Contact person: a lookup column in the "Site Pages" library that points to a list with the
 * columns "Title" (name) and an e-mail column. The form reads the e-mail address of the entry
 * selected on the page. If list or columns are missing, the form offers to create them
 * (for users with the permission to manage lists).
 */
export const FEEDBACK_CONFIG = {
  contact: {
    /**
     * Lookup column in the Site Pages library. internalName: letters, digits, "_" only.
     * Default, can be overridden per site (settings list, key settingsList.keys.contactColumn).
     */
    pageColumn: {
      internalName: 'Ansprechpartner',
      displayName: 'Ansprechpartner'
    },
    /**
     * List with the contact persons. url: relative to the site, always "Lists/<name>".
     * Default, can be overridden per site (settings list, key settingsList.keys.contactList).
     */
    list: {
      url: 'Lists/Ansprechpartner',
      title: 'Ansprechpartner'
    },
    /** E-mail column of the contact list (single line of text). The name comes from the "Title" column. */
    emailColumn: {
      internalName: 'Email',
      displayName: 'E-Mail'
    }
  },

  /**
   * Settings list of the site: one entry per setting, title = key, value in valueColumn (multiple lines of text).
   * url always "Lists/<name>". "Set up" in the settings view (gear icon, tab "Setup") creates list and column.
   * Visitors need read access (given with inherited permissions); everyone who may edit the list items can
   * change the settings - restrict the permissions of the list if that is not wanted.
   */
  settingsList: {
    url: 'Lists/FeedbackEinstellungen',
    title: 'Feedback-Einstellungen',
    valueColumn: {
      internalName: 'Wert',
      displayName: 'Wert'
    },
    /**
     * Titles of the entries. An empty or missing entry = the default of this file applies.
     * Categories and mail template are stored per language: the title gets the language code,
     * e.g. "Kategorien (de-de)", "Mail-Text (en-us)".
     */
    keys: {
      /** Name of the contact list in its URL ("Ansprechpartner" for "Lists/Ansprechpartner"). */
      contactList: 'Ansprechpartner-Liste',
      /** Internal name of the lookup column in Site Pages. */
      contactColumn: 'Ansprechpartner-Spalte',
      /** Per language: one category per line. Default: languages[...].categories. */
      categories: 'Kategorien',
      /** Per language: mail subject and body (both plain text). Default: languages[...].mailSubject/mailBody. */
      mailSubject: 'Mail-Betreff',
      mailBody: 'Mail-Text'
    }
  },

  /**
   * Shared mailbox of the site: an entry in the settings list (settingsList). Edited with "Design" permission
   * (gear icon in the dialog), works in SharePoint Online without custom scripts.
   */
  mailbox: {
    /** Title of the entry in the settings list. */
    key: 'Funktionspostfach'
  },

  /**
   * Notification mail, sent with SharePoint's REST API SP.Utilities.Utility.SendEmail.
   * Recipients: the contact person(s) of the form in To (required), the shared mailbox in Cc (optional,
   * together with cc below).
   */
  mail: {
    /** Fixed additional recipients, e.g. an archive mailbox. Invalid addresses are ignored. */
    cc: [] as string[],
    bcc: [] as string[],
    /**
     * Sender address. Empty = SharePoint's outgoing mail address.
     * SharePoint 2019: the SMTP server must accept this sender. SharePoint Online: usually ignored
     * (sender is no-reply@sharepointonline.com).
     */
    from: '',
    /** "Reply" in the mail program goes to the person who sent the report (header Reply-To). */
    replyToReporter: true
  },

  /**
   * Languages with their defaults. Each site can override categories and mail template per language
   * (settings list, settings view: tab "Form and e-mail"). The UI language picks the entry: exact match
   * ("de-de"), then the same language prefix ("de-at" -> "de-de"), otherwise defaultLanguage.
   *
   *  - label:       name of the language in the settings view
   *  - categories:  selectable categories of the form, in this order
   *  - mailSubject: plain text
   *  - mailBody:    plain text without HTML ("\n" = line break; a template with HTML tags is not used).
   *                 The mail is sent as HTML: the code escapes text and values, turns line breaks into <br>
   *                 and makes {{url}}, {{siteUrl}} and {{reporterEmail}} clickable links.
   *
   * Placeholders in subject and body:
   *   {{category}} {{subject}} {{description}} {{url}}
   *   {{reporterName}} {{reporterEmail}} {{siteTitle}} {{siteUrl}} {{date}} {{recipients}}
   */
  languages: {
    'de-de': {
      label: 'Deutsch',
      categories: ['Defekter Link', 'Fehlerhafter Inhalt', 'Veralteter Inhalt', 'Barrierefreiheit', 'Technisches Problem', 'Sonstiges'],
      mailSubject: '[Feedback] {{category}}: {{subject}}',
      mailBody: mailBody('Über den Feedback-Button wurde folgende Meldung gesendet:', {
        category: 'Kategorie', subject: 'Betreff', url: 'URL', description: 'Beschreibung',
        reporter: 'Gemeldet von', site: 'Website', date: 'Datum'
      })
    },
    'en-us': {
      label: 'English',
      categories: ['Broken link', 'Incorrect content', 'Outdated content', 'Accessibility', 'Technical problem', 'Other'],
      mailSubject: '[Feedback] {{category}}: {{subject}}',
      mailBody: mailBody('The following report was sent via the feedback button:', {
        category: 'Category', subject: 'Subject', url: 'URL', description: 'Description',
        reporter: 'Reported by', site: 'Site', date: 'Date'
      })
    },
    'fr-fr': {
      label: 'Français',
      categories: ['Lien rompu', 'Contenu erroné', 'Contenu obsolète', 'Accessibilité', 'Problème technique', 'Autre'],
      mailSubject: '[Feedback] {{category}} : {{subject}}',
      mailBody: mailBody('Le signalement suivant a été envoyé via le bouton de feedback :', {
        category: 'Catégorie', subject: 'Objet', url: 'URL', description: 'Description',
        reporter: 'Signalé par', site: 'Site', date: 'Date'
      })
    }
  } as { [language: string]: { label: string; categories: string[]; mailSubject: string; mailBody: string } },

  /** Language that applies if the UI language matches none of languages. */
  defaultLanguage: 'de-de'
};

/**
 * Default mail text: intro and one line per value, the description below. Only a helper for the templates
 * above - a template can just as well be written as a plain string ("\n" = line break).
 */
function mailBody(intro: string, labels: {
  category: string; subject: string; url: string; description: string; reporter: string; site: string; date: string;
}): string {
  return [
    intro,
    '',
    `${labels.category}: {{category}}`,
    `${labels.subject}: {{subject}}`,
    `${labels.url}: {{url}}`,
    '',
    `${labels.description}:`,
    '{{description}}',
    '',
    `${labels.reporter}: {{reporterName}} ({{reporterEmail}})`,
    `${labels.site}: {{siteTitle}} – {{siteUrl}}`,
    `${labels.date}: {{date}}`
  ].join('\n');
}
