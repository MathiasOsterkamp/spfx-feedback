/**
 * Checked on start; invalid values are logged to the browser console. Categories, mail templates and (with
 * contact.selectableInSettings) contact list/column can be overridden per site in the settings list; empty entry = this file.
 */
export const FEEDBACK_CONFIG = {
  /** Contact person: lookup column in "Site Pages" pointing to a list with "Title" (name) and an e-mail column. */
  contact: {
    /** true: contact list and column selectable per site in the settings view; false: read-only, this file applies. */
    selectableInSettings: false,
    /** Lookup column in Site Pages. internalName: letters, digits, "_" only. */
    pageColumn: {
      internalName: 'Ansprechpartner',
      displayName: 'Ansprechpartner'
    },
    /** Contact list. url: always "Lists/<name>". */
    list: {
      url: 'Lists/Ansprechpartner',
      title: 'Ansprechpartner'
    },
    /** E-mail column of the contact list (single line of text). */
    emailColumn: {
      internalName: 'Email',
      displayName: 'E-Mail'
    }
  },

  /**
   * Per-site settings list, created by "Set up". url: "Lists/<name>". Anyone who can edit its items can change
   * the settings - restrict its permissions if needed (visitors need read access).
   */
  settingsList: {
    url: 'Lists/FeedbackEinstellungen',
    title: 'Feedback-Einstellungen',
    valueColumn: {
      internalName: 'Wert',
      displayName: 'Wert'
    },
    /** Entry titles. Per-language entries get the language code, e.g. "Kategorien (de-de)". */
    keys: {
      /** Value: list name ("Ansprechpartner" for "Lists/Ansprechpartner"). Only with contact.selectableInSettings. */
      contactList: 'Ansprechpartner-Liste',
      /** Value: internal name of the lookup column. Only with contact.selectableInSettings. */
      contactColumn: 'Ansprechpartner-Spalte',
      /** Value: one category per line. */
      categories: 'Kategorien',
      mailSubject: 'Mail-Betreff',
      mailBody: 'Mail-Text'
    }
  },

  /** Shared mailbox (Cc): title of its entry in the settings list. */
  mailbox: {
    key: 'Funktionspostfach'
  },

  /** Notification mail (SP.Utilities.Utility.SendEmail). To: contact person(s), Cc: shared mailbox + cc. */
  mail: {
    /** Fixed additional recipients, e.g. an archive mailbox. Invalid addresses are ignored. */
    cc: [] as string[],
    bcc: [] as string[],
    /** Empty = SharePoint's sender. SharePoint 2019: the SMTP server must accept it; SharePoint Online: usually ignored. */
    from: '',
    /** Reply-To header = the reporter. */
    replyToReporter: true,
    /** "Send anonymously" checkbox: no reporter name/address and no Reply-To in the mail (SharePoint still sees the account). */
    allowAnonymous: true
  },

  /**
   * Per-language defaults. UI language picks: exact match, same prefix ("de-at" -> "de-de"), else defaultLanguage.
   * Templates are plain text, no HTML ("\n" = line break); {{url}}, {{siteUrl}} and {{reporterEmail}} become links.
   * Placeholders: {{category}} {{subject}} {{description}} {{url}} {{reporterName}} {{reporterEmail}}
   *   {{siteTitle}} {{siteUrl}} {{date}} {{recipients}}
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

  /** Fallback if the UI language matches none of languages. */
  defaultLanguage: 'de-de'
};

/** Builds the default templates; a template can just as well be a plain string. */
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
