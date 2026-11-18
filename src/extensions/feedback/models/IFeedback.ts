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
  /** Where the shared mailbox of the site is stored (config mailbox), code only. */
  mailbox: IMailboxSettings;
  /** Notification mail (config mail), template already chosen for the UI language. */
  mail: IMailSettings;
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

export type MailboxStorage = 'list' | 'propertyBag';

export interface IMailboxSettings {
  storage: MailboxStorage;
  /** Title of the entry in the settings list, or key in the property bag. Empty = no pre-filled mailbox. */
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
  /** HTML with {{placeholders}} */
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
