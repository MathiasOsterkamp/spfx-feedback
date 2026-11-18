/** A column: internal name (REST, CAML) and display name (used when it is created). */
export interface IColumnName {
  internalName: string;
  displayName: string;
}

/**
 * Checked configuration: config/feedbackConfig.ts (utils/sanitize.ts createSettings), possibly with the
 * values of the settings list applied (applyStoredValues).
 */
export interface IFeedbackSettings {
  /** Selectable categories, never empty. */
  categories: string[];
  /** Lookup column in the Site Pages library. Empty internal name = no contact (configuration invalid). */
  contactColumn: IColumnName;
  /** Contact list: URL relative to the site ("Lists/<name>") and title. */
  contactList: { url: string; title: string };
  /** E-mail column of the contact list (single line of text). */
  contactEmailColumn: IColumnName;
  /** Settings list of the site (config settingsList), code only. */
  settingsList: ISettingsListSettings;
  /** Entry of the shared mailbox in the settings list (config mailbox), code only. */
  mailbox: IMailboxSettings;
  /** Notification mail (config mail), template already chosen for the UI language. */
  mail: IMailSettings;
  /** Languages of the configuration with their defaults (code only), in the order of the configuration. */
  languages: ILanguageSettings[];
  /** Code of the language that applies to the UI language (one of languages). */
  language: string;
}

/** Defaults of one language (config languages). Categories and mail template can be overridden per site. */
export interface ILanguageSettings {
  /** Lower case, e.g. "de-de" */
  code: string;
  /** Name shown in the settings view, e.g. "Deutsch" */
  label: string;
  categories: string[];
  /** Plain text with {{placeholders}} */
  subject: string;
  /** Plain text with {{placeholders}} */
  body: string;
}

/** Titles of the entries in the settings list. */
export interface ISettingKeys {
  contactList: string;
  contactColumn: string;
  categories: string;
  mailSubject: string;
  mailBody: string;
}

export interface ISettingsListSettings {
  /** "Lists/<name>"; empty = configuration invalid, nothing is read or stored. */
  url: string;
  title: string;
  valueColumn: IColumnName;
  keys: ISettingKeys;
}

/** Shared mailbox: always an entry in the settings list. */
export interface IMailboxSettings {
  /** Title of the entry in the settings list. Empty = no pre-filled mailbox (also if the settings list is invalid). */
  key: string;
}

/** Entries of the settings list: title (key) -> value. */
export interface IStoredValues {
  [key: string]: string;
}

/** Settings of the current site, read when the dialog opens. */
export interface ISiteSettings {
  /** Code configuration with the values of the settings list applied. */
  settings: IFeedbackSettings;
  /** Entries of the settings list; undefined = could not be read (logged), the code configuration applies. */
  stored: IStoredValues | undefined;
}

export interface IMailSettings {
  /** Plain text with {{placeholders}} */
  subject: string;
  /** Plain text with {{placeholders}}, turned into HTML when the mail is built */
  body: string;
  cc: string[];
  bcc: string[];
  /** Empty = SharePoint's default sender */
  from: string;
  replyToReporter: boolean;
  /** UI language, used to format {{date}} */
  locale: string;
}

/** The current page, if it is an item of the Site Pages library (pages and news). */
export interface IPageInfo {
  listId: string;
  itemId: number;
  /** Current user may create lists and columns (Site Pages library and web). */
  canManageList: boolean;
  /** Current user may edit the page properties (approximated by the permission on the library). */
  canEditItems: boolean;
}

/**
 * Result of reading a default recipient from SharePoint:
 *  - ok: valid address(es) found
 *  - empty: column exists, but the page has no value
 *  - invalid: a value exists, but contains no valid address
 *  - missing: the column (contact) or the property bag key (mailbox) does not exist
 *  - wrongType: the column exists, but is no lookup column
 *  - lookupEmailMissing: the list the lookup column points to has no e-mail column
 *  - notConfigured: no valid column/key in the configuration
 *  - error: reading failed (no access, network, ...)
 */
export type RecipientStatus = 'ok' | 'empty' | 'invalid' | 'missing' | 'wrongType' | 'lookupEmailMissing' | 'notConfigured' | 'error';

export interface IRecipientSource {
  status: RecipientStatus;
  emails: string[];
  /** Contact only: ID of the list the lookup column points to (if the column exists). */
  lookupListId?: string;
}

/** Default recipients read from SharePoint. */
export interface IRecipientDefaults {
  contact: IRecipientSource;
  mailbox: IRecipientSource;
}

/**
 * State of one part the setup creates (settings view, tab "Setup"):
 *  - present: exists (and has the right type)
 *  - missing: does not exist, "Set up" creates it
 *  - needsUpdate: exists, but "Set up" changes it (settings list: visible -> hidden)
 *  - broken: the lookup column exists, but its list was deleted ("Set up" deletes and creates it anew)
 *  - wrongType: a column with this name exists, but it is no lookup column (cannot be fixed automatically)
 *  - notConfigured: the name in the code configuration is invalid (logged on start)
 */
export type SetupPartStatus = 'present' | 'missing' | 'needsUpdate' | 'broken' | 'wrongType' | 'notConfigured';

/** A list of the site: ID, title and URL relative to the server. */
export interface IListRef {
  id: string;
  title: string;
  serverRelativeUrl: string;
}

/** What exists of the setup in the current site (SetupService.check). */
export interface ISetupStatus {
  /** Settings list (config settingsList.url); hidden in the site contents, otherwise needsUpdate */
  settingsList: SetupPartStatus;
  /** Value column of the settings list */
  valueColumn: SetupPartStatus;
  /** One entry per setting in the settings list (created with the defaults of the code) */
  entries: SetupPartStatus;
  /** Titles of the entries that do not exist yet */
  missingEntries: string[];
  /**
   * Contact list: the target of the existing lookup column, otherwise the list set in the settings
   * (config contact.list or settings list).
   */
  contactList: SetupPartStatus;
  /** E-mail column in the list the lookup column points to - or, without lookup column, in the contact list */
  emailColumn: SetupPartStatus;
  /** Lookup column in the Site Pages library */
  pageColumn: SetupPartStatus;
  /** The existing lists, for links: settings list, contact list and the target of the lookup column. */
  lists: { settings?: IListRef; contact?: IListRef; lookup?: IListRef };
}


/** Content of the form. The address fields may contain several addresses, separated by ";" or ",". */
export interface IFeedbackForm {
  category: string;
  subject: string;
  url: string;
  description: string;
  contactEmail: string;
  mailboxEmail: string;
}

export type FeedbackField = keyof IFeedbackForm;

/** Validation messages per field; a missing entry means the field is valid. */
export type FeedbackErrors = { [K in FeedbackField]?: string };

/** Who sends the report and from where (part of the mail). */
export interface IFeedbackOrigin {
  userName: string;
  userEmail: string;
  webTitle: string;
  webUrl: string;
}
