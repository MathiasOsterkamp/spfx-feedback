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
  /** Selectable categories of the form, in this order. */
  categories: string[];
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

  setupTitle: string;
  setupVisibility: string;
  setupContactInvalid: string;
  setupEditPageProperties: string;
  setupMailboxInvalid: string;
  setupMailboxLabel: string;
  setupSaved: string;
  setupOpen: string;
  setupBack: string;
  setupIncomplete: string;
  setupContactHeading: string;
  setupContactMissing: string;
  setupContactSetup: string;
  setupContactDone: string;
  setupPageContactNone: string;
  setupNotConfigured: string;
  setupMailboxHeading: string;
  setupMailboxNeedsOwner: string;
  setupColumnLabel: string;
  setupListLabel: string;
  setupPageContactLabel: string;
  setupStatusPresent: string;
  setupStatusMissing: string;
  setupStatusWrongType: string;
  setupStatusEmailMissing: string;
  setupStatusUnknown: string;
  setupMailboxDescription: string;
  setupMailboxDescriptionList: string;
  setupSave: string;
  setupFailed: string;
  setupAccessDenied: string;
  setupReadFailed: string;

  /** Settings list (setup view). Templates with {0}: see loc/template.ts. */
  setupSettingsReadFailed: string;
  setupSettingsStored: string;
  setupContactListField: string;
  setupContactListDescription: string;
  setupContactColumnField: string;
  setupContactColumnDescription: string;
  setupFormHeading: string;
  setupCategoriesLabel: string;
  setupCategoriesDescription: string;
  setupMailHeading: string;
  setupMailSubjectLabel: string;
  setupMailBodyLabel: string;
  setupMailBodyPlaceholder: string;
  setupMailDescription: string;
  setupInsertDefault: string;
  setupErrorListName: string;
  setupErrorColumnName: string;
  setupErrorCategories: string;
  setupErrorPlaceholder: string;
}

declare module 'FeedbackApplicationCustomizerStrings' {
  const strings: IFeedbackApplicationCustomizerStrings;
  export = strings;
}
