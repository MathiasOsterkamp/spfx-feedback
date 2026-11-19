export interface IColumnName {
  internalName: string;
  displayName: string;
}

export interface IFeedbackSettings {
  /** Never empty. */
  categories: string[];
  /** Empty internalName = contact disabled (invalid configuration). */
  contactColumn: IColumnName;
  /** url: "Lists/<name>" */
  contactList: { url: string; title: string };
  contactEmailColumn: IColumnName;
  settingsList: ISettingsListSettings;
  mailbox: IMailboxSettings;
  mail: IMailSettings;
  languages: ILanguageSettings[];
  language: string;
  /** Configured defaultLanguage, or the first language if that is not configured. */
  defaultLanguage: string;
}

export interface ILanguageSettings {
  /** Lower case, e.g. "de-de" */
  code: string;
  label: string;
  categories: string[];
  subject: string;
  body: string;
}

/** Settings list entry titles. Empty = setting not changeable per site. */
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
  /** Settings list entry title. Empty = no pre-filled mailbox (also if the settings list is invalid). */
  key: string;
}

export interface IStoredValues {
  [key: string]: string;
}

export interface ISiteSettings {
  settings: IFeedbackSettings;
  /** undefined = could not be read (logged), the code configuration applies. */
  stored: IStoredValues | undefined;
}

export interface IMailSettings {
  subject: string;
  /** Plain text with {{placeholders}}, turned into HTML when the mail is built. */
  body: string;
  cc: string[];
  bcc: string[];
  /** Empty = SharePoint's default sender */
  from: string;
  replyToReporter: boolean;
  allowAnonymous: boolean;
  /** Used to format {{date}} */
  locale: string;
}

export interface IPageInfo {
  listId: string;
  itemId: number;
  canManageList: boolean;
  /** Approximated by the permission on the library. */
  canEditItems: boolean;
}

/** invalid: no valid address; wrongType: no lookup column; lookupEmailMissing: lookup target list has no e-mail column. */
export type RecipientStatus = 'ok' | 'empty' | 'invalid' | 'missing' | 'wrongType' | 'lookupEmailMissing' | 'notConfigured' | 'error';

export interface IRecipientSource {
  status: RecipientStatus;
  emails: string[];
}

export interface IRecipientDefaults {
  contact: IRecipientSource;
  mailbox: IRecipientSource;
}

/** needsUpdate: settings list visible -> hidden; broken: lookup to a deleted list (recreated); wrongType: no lookup (not auto-fixable). */
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

export type FeedbackErrors = { [K in FeedbackField]?: string };

export interface IFeedbackOrigin {
  userName: string;
  userEmail: string;
  webTitle: string;
  webUrl: string;
}
