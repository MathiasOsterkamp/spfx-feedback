export interface IColumnName {
  internalName: string;
  displayName: string;
}

export interface IFeedbackSettings {
  categories: string[];
  contactColumn: IColumnName;
  contactList: { url: string; title: string };
  contactEmailColumn: IColumnName;
  settingsList: ISettingsListSettings;
  mailbox: IMailboxSettings;
  mail: IMailSettings;
  languages: ILanguageSettings[];
  language: string;
  defaultLanguage: string;
}

export interface ILanguageSettings {
  code: string;
  label: string;
  categories: string[];
  subject: string;
  body: string;
}

export interface ISettingKeys {
  contactList: string;
  contactColumn: string;
  categories: string;
  mailSubject: string;
  mailBody: string;
}

export interface ISettingsListSettings {
  url: string;
  title: string;
  valueColumn: IColumnName;
  keys: ISettingKeys;
}

export interface IMailboxSettings {
  key: string;
}

export interface IStoredValues {
  [key: string]: string;
}

export interface ISiteSettings {
  settings: IFeedbackSettings;
  stored: IStoredValues | undefined;
}

export interface IMailSettings {
  subject: string;
  body: string;
  cc: string[];
  bcc: string[];
  from: string;
  replyToReporter: boolean;
  allowAnonymous: boolean;
  locale: string;
}

export interface IPageInfo {
  listId: string;
  itemId: number;
  canManageList: boolean;
  canEditItems: boolean;
}

export type RecipientStatus = 'ok' | 'empty' | 'invalid' | 'missing' | 'wrongType' | 'lookupEmailMissing' | 'notConfigured' | 'error';

export interface IRecipientSource {
  status: RecipientStatus;
  emails: string[];
}

export interface IRecipientDefaults {
  contact: IRecipientSource;
  mailbox: IRecipientSource;
}

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
  contactList: SetupPartStatus;
  emailColumn: SetupPartStatus;
  pageColumn: SetupPartStatus;
  lookupList?: IListRef;
}

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
