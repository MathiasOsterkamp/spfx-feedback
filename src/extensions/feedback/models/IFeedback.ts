export interface IColumnName {
  internalName: string;
  displayName: string;
}

/** Checked code configuration (createSettings), possibly with the settings list applied (applyStoredValues). */
export interface IFeedbackSettings {
  /** Never empty. */
  categories: string[];
  /** Lookup column in Site Pages. Empty internal name = contact disabled (configuration invalid). */
  contactColumn: IColumnName;
  /** url: "Lists/<name>" */
  contactList: { url: string; title: string };
  contactEmailColumn: IColumnName;
  settingsList: ISettingsListSettings;
  mailbox: IMailboxSettings;
  /** Template already chosen for the UI language. */
  mail: IMailSettings;
  languages: ILanguageSettings[];
  /** Code of the language that applies to the UI language (one of languages). */
  language: string;
}

export interface ILanguageSettings {
  /** Lower case, e.g. "de-de" */
  code: string;
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

export interface IMailboxSettings {
  /** Title of the settings list entry. Empty = no pre-filled mailbox (also if the settings list is invalid). */
  key: string;
}

/** Entries of the settings list: title (key) -> value. */
export interface IStoredValues {
  [key: string]: string;
}

export interface ISiteSettings {
  settings: IFeedbackSettings;
  /** undefined = could not be read (logged), the code configuration applies. */
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

/** The current page, if it is an item of the Site Pages library. */
export interface IPageInfo {
  listId: string;
  itemId: number;
  /** May create lists and columns. */
  canManageList: boolean;
  /** Approximated by the permission on the library. */
  canEditItems: boolean;
}

/**
 * empty: page has no value; invalid: value without valid address; missing: column/key does not exist;
 * wrongType: no lookup column; lookupEmailMissing: lookup target list has no e-mail column;
 * notConfigured: no valid column/key in the configuration; error: reading failed.
 */
export type RecipientStatus = 'ok' | 'empty' | 'invalid' | 'missing' | 'wrongType' | 'lookupEmailMissing' | 'notConfigured' | 'error';

export interface IRecipientSource {
  status: RecipientStatus;
  emails: string[];
}

export interface IRecipientDefaults {
  contact: IRecipientSource;
  mailbox: IRecipientSource;
}

/**
 * needsUpdate: settings list visible -> hidden; broken: lookup column whose list was deleted (recreated);
 * wrongType: column exists but is no lookup (not fixable automatically); notConfigured: invalid name in the code configuration.
 */
export type SetupPartStatus = 'present' | 'missing' | 'needsUpdate' | 'broken' | 'wrongType' | 'notConfigured';

export interface IListRef {
  id: string;
  title: string;
  serverRelativeUrl: string;
}

export interface ISetupStatus {
  settingsList: SetupPartStatus;
  valueColumn: SetupPartStatus;
  entries: SetupPartStatus;
  /** Target of the existing lookup column, otherwise the configured contact list. */
  contactList: SetupPartStatus;
  /** In the lookup target list - or, without lookup column, in the contact list. */
  emailColumn: SetupPartStatus;
  pageColumn: SetupPartStatus;
  /** Target list of the existing lookup column */
  lookupList?: IListRef;
}


/** The address fields may contain several addresses, separated by ";" or ",". */
export interface IFeedbackForm {
  category: string;
  subject: string;
  url: string;
  description: string;
  contactEmail: string;
  mailboxEmail: string;
}

export type FeedbackField = keyof IFeedbackForm;

/** A missing entry means the field is valid. */
export type FeedbackErrors = { [K in FeedbackField]?: string };

export interface IFeedbackOrigin {
  userName: string;
  userEmail: string;
  webTitle: string;
  webUrl: string;
}
