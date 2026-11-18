/**
 * CONFIGURATION OF THE FEEDBACK BUTTON
 *
 * Rolled out with the package. The values are checked on start; invalid names are reported in the
 * browser console.
 *
 * Some settings can be overridden per site in the settings list (see settingsList). They are edited in
 * the setup view of the dialog (gear icon): contact list and lookup column, categories, mail subject and
 * mail text. An empty or missing entry there means: the value from this file applies.
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
   * url always "Lists/<name>". The setup view (gear icon) creates list and column when it saves for the first time.
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
    /** Titles of the entries. An empty or missing entry = the default of this file applies. */
    keys: {
      /** Name of the contact list in its URL ("Ansprechpartner" for "Lists/Ansprechpartner"). */
      contactList: 'Ansprechpartner-Liste',
      /** Internal name of the lookup column in Site Pages. */
      contactColumn: 'Ansprechpartner-Spalte',
      /** One category per line, for all languages. Default: "categories" of the language files. */
      categories: 'Kategorien',
      /** Mail subject (plain text) and body (HTML) for all languages. Default: mail.templates. */
      mailSubject: 'Mail-Betreff',
      mailBody: 'Mail-Text'
    }
  },

  /**
   * Shared mailbox of the site. Where it is stored is fixed here:
   *  - 'list':        entry in the settings list (settingsList). Can be edited with "Design" permission
   *                   (gear icon in the dialog) and works in SharePoint Online without custom scripts.
   *  - 'propertyBag': property bag of the site. Only site owners can change it; in SharePoint Online
   *                   custom scripts must be allowed for the site.
   */
  mailbox: {
    storage: 'list' as 'list' | 'propertyBag',
    /** Title of the entry in the settings list, or key in the property bag. */
    key: 'Funktionspostfach'
  },

  // The default categories are texts of the language files (loc/de-de.js, en-us.js, fr-fr.js: "categories").

  /**
   * Notification mail, sent with SharePoint's REST API SP.Utilities.Utility.SendEmail.
   * Recipients (To) are the contact person(s) and the shared mailbox from the form.
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
    replyToReporter: true,
    /**
     * Default subject (plain text) and body (HTML) per UI language; the settings list can override them for
     * all languages. Unknown languages use the entry with the same language prefix (e.g. "de-at" -> "de-de"),
     * otherwise defaultLanguage.
     *
     * Placeholders (values are always HTML-escaped; line breaks of the description become <br>):
     *   {{category}} {{subject}} {{description}} {{url}}
     *   {{reporterName}} {{reporterEmail}} {{siteTitle}} {{siteUrl}} {{date}} {{recipients}}
     *
     * Mail programs ignore <style> blocks and external CSS in many cases: use inline styles.
     */
    templates: {
      'de-de': {
        subject: '[Feedback] {{category}}: {{subject}}',
        body: mailBody('Über den Feedback-Button wurde folgende Meldung gesendet:', {
          category: 'Kategorie', subject: 'Betreff', url: 'URL', description: 'Beschreibung',
          reporter: 'Gemeldet von', site: 'Website', date: 'Datum'
        })
      },
      'en-us': {
        subject: '[Feedback] {{category}}: {{subject}}',
        body: mailBody('The following report was sent via the feedback button:', {
          category: 'Category', subject: 'Subject', url: 'URL', description: 'Description',
          reporter: 'Reported by', site: 'Site', date: 'Date'
        })
      },
      'fr-fr': {
        subject: '[Feedback] {{category}} : {{subject}}',
        body: mailBody('Le signalement suivant a été envoyé via le bouton de feedback :', {
          category: 'Catégorie', subject: 'Objet', url: 'URL', description: 'Description',
          reporter: 'Signalé par', site: 'Site', date: 'Date'
        })
      }
    } as { [language: string]: { subject: string; body: string } }
  },

  /** Language whose mail template applies if the UI language has none. */
  defaultLanguage: 'de-de'
};

/**
 * Default mail layout: intro and a table with all values. Only a helper for the templates above -
 * a template can just as well be written as a complete HTML string.
 */
function mailBody(intro: string, labels: {
  category: string; subject: string; url: string; description: string; reporter: string; site: string; date: string;
}): string {
  const th = 'padding:6px 12px;border:1px solid #c8c6c4;vertical-align:top;text-align:left;background:#f3f2f1;white-space:nowrap;';
  const td = 'padding:6px 12px;border:1px solid #c8c6c4;vertical-align:top;text-align:left;';
  const row = (label: string, value: string): string => `<tr><th style="${th}">${label}</th><td style="${td}">${value}</td></tr>`;
  return '<div style="font-family:\'Segoe UI\',Arial,sans-serif;font-size:14px;color:#323130;">'
    + `<p>${intro}</p>`
    + '<table style="border-collapse:collapse;">'
    + row(labels.category, '{{category}}')
    + row(labels.subject, '{{subject}}')
    + row(labels.url, '<a href="{{url}}">{{url}}</a>')
    + row(labels.description, '{{description}}')
    + row(labels.reporter, '{{reporterName}} &lt;<a href="mailto:{{reporterEmail}}">{{reporterEmail}}</a>&gt;')
    + row(labels.site, '<a href="{{siteUrl}}">{{siteTitle}}</a>')
    + row(labels.date, '{{date}}')
    + '</table>'
    + '</div>';
}
