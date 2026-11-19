export const FEEDBACK_CONFIG = {
  contact: {
    selectableInSettings: false,
    // internalName: letters, digits and "_" only
    pageColumn: {
      internalName: 'Ansprechpartner',
      displayName: 'Ansprechpartner'
    },
    // url: always "Lists/<name>"
    list: {
      url: 'Lists/Ansprechpartner',
      title: 'Ansprechpartner'
    },
    emailColumn: {
      internalName: 'Email',
      displayName: 'E-Mail'
    }
  },

  // Anyone who can edit its items can change the settings: restrict the list permissions (visitors need read access).
  settingsList: {
    url: 'Lists/FeedbackEinstellungen',
    title: 'Feedback-Einstellungen',
    valueColumn: {
      internalName: 'Wert',
      displayName: 'Wert'
    },
    keys: {
      contactList: 'Ansprechpartner-Liste',
      contactColumn: 'Ansprechpartner-Spalte',
      categories: 'Kategorien',
      mailSubject: 'Mail-Betreff',
      mailBody: 'Mail-Text'
    }
  },

  mailbox: {
    key: 'Funktionspostfach'
  },

  mail: {
    cc: [] as string[],
    bcc: [] as string[],
    from: '',
    replyToReporter: true,
    allowAnonymous: true
  },

  // Templates: plain text, no HTML. Placeholders: {{category}} {{subject}} {{description}} {{url}} {{reporterName}}
  // {{reporterEmail}} {{siteTitle}} {{siteUrl}} {{date}} {{recipients}}
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

  defaultLanguage: 'de-de'
};

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
