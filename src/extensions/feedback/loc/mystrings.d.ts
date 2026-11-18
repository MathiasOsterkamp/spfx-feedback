declare interface IFeedbackApplicationCustomizerStrings {
  /** Plural rule of the language for formatText: non-empty = 0 counts as singular (French). */
  zeroIsSingular: string;

  buttonLabel: string;
  dialogTitle: string;
  close: string;
  send: string;
  sending: string;
  cancel: string;

  sectionRequest: string;
  sectionResponsibility: string;
  categoryLabel: string;
  categoryPlaceholder: string;
  subjectLabel: string;
  subjectPlaceholder: string;
  urlLabel: string;
  urlPlaceholder: string;
  descriptionLabel: string;
  descriptionPlaceholder: string;
  contactLabel: string;
  contactPlaceholder: string;
  mailboxLabel: string;
  mailboxPlaceholder: string;
  responsibilityHint: string;
  loadingRecipients: string;
  contactMissing: string;
  mailboxMissing: string;

  errorRequired: string;
  errorCategory: string;
  errorUrl: string;
  errorEmail: string;
  errorEmailDomain: string;
  errorRecipient: string;
  errorTooManyEmails: string;
  errorTooLong: string;
  errorForm: string;
  sendFailed: string;
  sendSucceeded: string;

  /** Preview of the mail before sending. */
  previewSend: string;
  previewHeading: string;
  previewHint: string;
  previewTo: string;
  previewCc: string;
  previewReplyTo: string;
  previewSubject: string;
  /** Shown for an empty address line (e.g. no Cc). */
  previewNone: string;
  previewBack: string;

  /** Settings view (gear icon). Templates with {0}: see loc/template.ts. */
  settingsOpen: string;
  settingsTitle: string;
  settingsBack: string;
  settingsSave: string;
  settingsSaved: string;
  settingsSaveFailed: string;
  settingsAccessDenied: string;
  settingsUnsaved: string;
  settingsDiscard: string;
  settingsKeepEditing: string;
  settingsVisibility: string;
  settingsReadFailed: string;
  settingsNotSetUp: string;
  settingsGoToSetup: string;
  settingsNoPermission: string;
  settingsTabRecipients: string;
  settingsTabTexts: string;
  settingsTabSetup: string;
  /** Screen reader text of a tab with warning icon, {0} = tab name. */
  settingsTabWarning: string;
  settingsContactHeading: string;
  settingsContactListLabel: string;
  settingsContactListDescription: string;
  settingsContactListManual: string;
  settingsOptionDefault: string;
  settingsContactListLoading: string;
  settingsOptionNotFound: string;
  settingsContactListLookup: string;
  settingsContactColumnLabel: string;
  settingsContactColumnDescription: string;
  settingsContactColumnManual: string;
  settingsContactColumnLoading: string;
  settingsPageContactLabel: string;
  settingsPageContactNone: string;
  settingsPageContactInvalid: string;
  settingsPageContactUnavailable: string;
  settingsEditPageProperties: string;
  settingsMailboxHeading: string;
  settingsMailboxLabel: string;
  settingsMailboxDescriptionList: string;
  settingsMailboxNotConfigured: string;
  settingsLanguageLabel: string;
  settingsLanguageCustomized: string;
  settingsLanguageErrors: string;
  settingsFormHeading: string;
  settingsCategoriesLabel: string;
  settingsCategoriesDescription: string;
  settingsMailHeading: string;
  settingsMailSubjectLabel: string;
  settingsMailBodyLabel: string;
  settingsMailBodyPlaceholder: string;
  settingsMailDescription: string;
  settingsInsertDefault: string;
  settingsUseDefault: string;
  settingsSetupIntro: string;
  settingsSetupSettingsList: string;
  settingsSetupContactList: string;
  settingsSetupEmailColumn: string;
  settingsSetupPageColumn: string;
  settingsSetupUnknownList: string;
  settingsSetupLookupTarget: string;
  settingsStatusPresent: string;
  settingsStatusMissing: string;
  settingsStatusVisible: string;
  settingsStatusIncomplete: string;
  settingsStatusBroken: string;
  settingsStatusWrongType: string;
  settingsStatusNotConfigured: string;
  settingsSetupRun: string;
  settingsSetupRefresh: string;
  settingsSetupComplete: string;
  settingsSetupDone: string;
  settingsSetupFailed: string;
  settingsSetupCheckFailed: string;
  settingsSetupNeedsPermission: string;
  settingsSetupSaveFirst: string;
  settingsSetupNothingToDo: string;

  setupErrorListName: string;
  setupErrorColumnName: string;
  setupErrorHtml: string;
  setupErrorCategories: string;
  setupErrorPlaceholder: string;
}

declare module 'FeedbackApplicationCustomizerStrings' {
  const strings: IFeedbackApplicationCustomizerStrings;
  export = strings;
}
