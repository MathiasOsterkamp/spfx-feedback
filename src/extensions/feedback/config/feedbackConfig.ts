/**
 * Configuration of the feedback button, rolled out with the package. Checked on start; invalid values are
 * logged to the browser console. Contact list, lookup column, categories and mail template can be overridden
 * per site in the settings list (gear icon in the dialog); an empty entry there = the value from this file.
 *
 * Contact person: a lookup column in "Site Pages" pointing to a list with "Title" (name) and an e-mail column.
 */
export const FEEDBACK_CONFIG = {
  contact: {
    /** Lookup column in Site Pages. internalName: letters, digits, "_" only. Overridable per site. */
    pageColumn: {
      internalName: 'Ansprechpartner',
      displayName: 'Ansprechpartner'
    },
    /** Contact list. url: always "Lists/<name>". Overridable per site. */
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
   * Settings list of the site: one entry per setting, title = key, value in valueColumn. url: "Lists/<name>".
   * Created by "Set up" in the settings view. Everyone who may edit its items can change the settings -
   * restrict the list permissions if that is not wanted (visitors need read access).
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
      /** Value: list name in its URL ("Ansprechpartner" for "Lists/Ansprechpartner"). */
      contactList: 'Ansprechpartner-Liste',
      /** Value: internal name of the lookup column. */
      contactColumn: 'Ansprechpartner-Spalte',
      /** Per language, one category per line. */
      categories: 'Kategorien',
      /** Per language, plain text. */
      mailSubject: 'Mail-Betreff',
      mailBody: 'Mail-Text'
    }
  },

  /** Shared mailbox (Cc): title of its entry in the settings list. Works in SharePoint Online without custom scripts. */
  mailbox: {
    key: 'Funktionspostfach'
  },

  /** Notification mail (SP.Utilities.Utility.SendEmail). To: contact person(s), Cc: shared mailbox + cc. */
  mail: {
    /** Fixed additional recipients, e.g. an archive mailbox. Invalid addresses are ignored. */
    cc: [] as string[],
    bcc: [] as string[],
    /**
     * Empty = SharePoint's outgoing address. SharePoint 2019: the SMTP server must accept this sender.
     * SharePoint Online: usually ignored (no-reply@sharepointonline.com).
     */
    from: '',
    /** Reply-To header = the reporter. */
    replyToReporter: true
  },

  /**
   * Per-language defaults (overridable per site). The UI language picks the entry: exact match, then the
   * same prefix ("de-at" -> "de-de"), otherwise defaultLanguage.
   * mailSubject/mailBody: plain text, no HTML ("\n" = line break). The mail is sent as HTML: text and values
   * are escaped, {{url}}, {{siteUrl}} and {{reporterEmail}} become links.
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

/** Helper for the default templates above; a template can just as well be a plain string. */
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
