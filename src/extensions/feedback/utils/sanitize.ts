import { logError, logWarning } from './logging';
import { asRecord } from './untyped';
import { checkEmail } from './validation';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackSettings, ILanguageSettings, IMailboxSettings, IMailSettings, ISettingsListSettings, IStoredValues } from '../models/IFeedback';

const COLOR_PATTERN = /^(#[0-9a-f]{3,8}|rgba?\(\s*[\d.\s,%]+\)|hsla?\(\s*[\d.\s,%deg]+\)|[a-z]{3,30})$/i;

/** Internal field names (special characters encoded as _xHHHH_). Inserted into REST URLs, so nothing else is allowed. */
export const FIELD_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;

/** List name in "Lists/<name>"; becomes part of REST URLs. */
export const LIST_NAME = /^[A-Za-z0-9_-]{1,128}$/;

/** Prevents CSS injection via color values. */
export function safeColor(value: unknown, fallback: string): string {
  const v = typeof value === 'string' ? value.trim() : '';
  return COLOR_PATTERN.test(v) ? v : fallback;
}

export function singleLine(value: string): string {
  // eslint-disable-next-line no-control-regex -- control characters are exactly what is removed here
  return value.replace(/[\u0000-\u001f\u007f]+/g, ' ').trim();
}

export interface IFeedbackConfig {
  contact: {
    pageColumn: { internalName: string; displayName: string };
    list: { url: string; title: string };
    emailColumn: { internalName: string; displayName: string };
  };
  settingsList: {
    url: string;
    title: string;
    valueColumn: { internalName: string; displayName: string };
    keys: { contactList: string; contactColumn: string; categories: string; mailSubject: string; mailBody: string };
  };
  mailbox: {
    key: string;
  };
  mail: {
    cc: string[];
    bcc: string[];
    from: string;
    replyToReporter: boolean;
  };
  languages: { [language: string]: { label: string; categories: string[]; mailSubject: string; mailBody: string } };
  defaultLanguage: string;
}

/** Lower case: "de", "de-de", "zh-hant-tw". */
const LANGUAGE_CODE = /^[a-z]{2,3}(-[a-z0-9]{2,8})*$/;

/** "Kategorien (de-de)". Empty key = no entry. */
export function languageKey(key: string, language: string): string {
  return key ? `${key} (${language})` : '';
}

export function settingEntryKeys(settings: IFeedbackSettings): string[] {
  const keys = settings.settingsList.keys;
  return settings.languages
    .reduce((all: string[], language: ILanguageSettings) => all.concat(
      [keys.categories, keys.mailSubject, keys.mailBody].map((key: string) => languageKey(key, language.code))
    ), [keys.contactList, keys.contactColumn, settings.mailbox.key])
    .filter((key: string) => !!key);
}

/** All settings list entries with the code defaults, as "Set up" creates them. */
export function defaultSettingEntries(base: IFeedbackSettings): IStoredValues {
  const keys = base.settingsList.keys;
  const entries: IStoredValues = {};
  const add = (key: string, value: string): void => {
    if (key) { entries[key] = value; }
  };
  add(keys.contactList, base.contactList.url.substring('Lists/'.length));
  add(keys.contactColumn, base.contactColumn.internalName);
  add(base.mailbox.key, '');
  base.languages.forEach((language: ILanguageSettings) => {
    add(languageKey(keys.categories, language.code), language.categories.join('\n'));
    add(languageKey(keys.mailSubject, language.code), language.subject);
    add(languageKey(keys.mailBody, language.code), language.body);
  });
  return entries;
}

const FALLBACK_TEMPLATE = { subject: '[Feedback] {{subject}}', body: '{{category}}: {{subject}}\n{{url}}\n\n{{description}}\n\n{{reporterName}} {{reporterEmail}}' };

/** HTML tags: mail templates are plain text, tags would appear literally in the mail. */
const HTML_TAG = /<\/?[a-z!][^>]*>/i;

export function containsHtml(template: string): boolean {
  return HTML_TAG.test(template);
}

const LIST_URL = /^Lists\/[A-Za-z0-9_-]{1,128}$/;

/**
 * Checks config/feedbackConfig.ts and picks the UI language. Invalid names are logged and switch the
 * affected feature off instead of sending broken requests.
 */
export function createSettings(config: IFeedbackConfig, uiLanguage: string, policy: ISecurityPolicy): IFeedbackSettings {
  const contact = config.contact;
  const columnName = safeFieldName(contact.pageColumn.internalName, 'contact.pageColumn.internalName');
  const emailName = safeFieldName(contact.emailColumn.internalName, 'contact.emailColumn.internalName');
  const listUrl = LIST_URL.test(contact.list.url) ? contact.list.url : '';
  if (!listUrl) {
    logError('Invalid configuration, contact disabled', { setting: 'contact.list.url', expected: 'Lists/<name>' });
  }
  const contactValid = !!columnName && !!emailName && !!listUrl;
  const settingsList = createSettingsList(config);
  const languages = createLanguages(config, policy);
  const language = pickLanguage(languages, uiLanguage, config.defaultLanguage);
  return {
    categories: language.categories,
    contactColumn: {
      internalName: contactValid ? columnName : '',
      displayName: singleLine(contact.pageColumn.displayName || '') || columnName
    },
    contactList: { url: listUrl, title: singleLine(contact.list.title || '') || listUrl.substr(6) },
    contactEmailColumn: { internalName: emailName, displayName: singleLine(contact.emailColumn.displayName || '') || emailName },
    settingsList,
    mailbox: createMailboxSettings(config, settingsList),
    mail: createMailSettings(config, language, uiLanguage, policy),
    languages,
    language: language.code
  };
}

/** Missing categories/template get a minimal one (logged); without any usable language a minimal English entry applies. */
function createLanguages(config: IFeedbackConfig, policy: ISecurityPolicy): ILanguageSettings[] {
  const entries = asRecord(config.languages);
  const result: ILanguageSettings[] = [];
  Object.keys(entries).forEach((name: string) => {
    const code = name.trim().toLowerCase();
    const setting = 'languages.' + name;
    if (!LANGUAGE_CODE.test(code) || result.some((language: ILanguageSettings) => language.code === code)) {
      logError('Invalid or duplicate language code in the configuration, ignored', { setting });
      return;
    }
    const entry = asRecord(entries[name]);
    const categories = safeCategories(entry.categories, policy);
    if (!categories) {
      logError('No usable categories in the configuration', { setting: setting + '.categories' });
    }
    const subject = typeof entry.mailSubject === 'string' ? singleLine(entry.mailSubject) : '';
    const body = typeof entry.mailBody === 'string' ? entry.mailBody.trim() : '';
    const templateUsable = !!subject && !!body && !containsHtml(body);
    if (!templateUsable) {
      logError('No usable mail template in the configuration (empty or contains HTML), using a minimal one', { setting });
    }
    result.push({
      code,
      label: (typeof entry.label === 'string' ? singleLine(entry.label) : '') || code,
      categories: categories || ['-'],
      subject: templateUsable ? subject : FALLBACK_TEMPLATE.subject,
      body: templateUsable ? body : FALLBACK_TEMPLATE.body
    });
  });
  if (result.length === 0) {
    logError('No usable language in the configuration', { setting: 'languages' });
    result.push({ code: 'en-us', label: 'English', categories: ['-'], ...FALLBACK_TEMPLATE });
  }
  return result;
}

function pickLanguage(languages: ILanguageSettings[], uiLanguage: string, defaultLanguage: string): ILanguageSettings {
  const byCode: { [code: string]: ILanguageSettings } = {};
  languages.forEach((language: ILanguageSettings) => { byCode[language.code] = language; });
  const fallback = typeof defaultLanguage === 'string' ? byCode[defaultLanguage.trim().toLowerCase()] : undefined;
  return pickByLanguage(byCode, uiLanguage) || fallback || languages[0];
}

function createSettingsList(config: IFeedbackConfig): ISettingsListSettings {
  const list = config.settingsList;
  const keys = list.keys;
  let url = LIST_URL.test(list.url) ? list.url : '';
  const column = safeFieldName(list.valueColumn.internalName, 'settingsList.valueColumn.internalName');
  if (!url) {
    logError('Invalid configuration, settings list disabled', { setting: 'settingsList.url', expected: 'Lists/<name>' });
  }
  if (!column) {
    url = '';
  }
  return {
    url,
    title: singleLine(list.title || '') || url.substr(6),
    valueColumn: { internalName: column, displayName: singleLine(list.valueColumn.displayName || '') || column },
    keys: {
      contactList: safeEntryKey(keys.contactList),
      contactColumn: safeEntryKey(keys.contactColumn),
      categories: safeEntryKey(keys.categories),
      mailSubject: safeEntryKey(keys.mailSubject),
      mailBody: safeEntryKey(keys.mailBody)
    }
  };
}

function createMailboxSettings(config: IFeedbackConfig, settingsList: ISettingsListSettings): IMailboxSettings {
  const key = safeEntryKey(config.mailbox.key);
  // An invalid settings list is already logged: no mailbox from there.
  return { key: settingsList.url ? key : '' };
}

/**
 * Empty entries keep the code value. Invalid entries are logged and ignored - the settings view does not
 * store them, but the list can be edited directly.
 */
export function applyStoredValues(base: IFeedbackSettings, stored: IStoredValues, policy: ISecurityPolicy): IFeedbackSettings {
  const keys = base.settingsList.keys;
  const value = (key: string): string => (key && typeof stored[key] === 'string' ? stored[key].trim() : '');
  const ignored = (key: string): void => logError('Invalid value in the settings list, ignored', { key });
  const result: IFeedbackSettings = { ...base, mail: { ...base.mail } };
  const categoriesKey = languageKey(keys.categories, base.language);
  const subjectKey = languageKey(keys.mailSubject, base.language);
  const bodyKey = languageKey(keys.mailBody, base.language);

  const categories = value(categoriesKey);
  if (categories) {
    const parsed = safeCategories(splitLines(categories), policy);
    if (parsed) { result.categories = parsed; } else { ignored(categoriesKey); }
  }

  const listName = value(keys.contactList);
  if (listName) {
    const url = 'Lists/' + listName;
    if (!LIST_NAME.test(listName)) {
      ignored(keys.contactList);
    } else if (url.toLowerCase() !== base.contactList.url.toLowerCase()) {
      result.contactList = { url, title: listName };
    }
  }

  // Empty internal name in base = invalid code configuration, the contact stays off.
  const columnName = value(keys.contactColumn);
  if (columnName && base.contactColumn.internalName) {
    if (!FIELD_NAME.test(columnName)) {
      ignored(keys.contactColumn);
    } else if (columnName !== base.contactColumn.internalName) {
      result.contactColumn = { internalName: columnName, displayName: columnName };
    }
  }

  const subject = value(subjectKey);
  if (subject) {
    if (subject.length <= policy.maxTextLength) { result.mail.subject = singleLine(subject); } else { ignored(subjectKey); }
  }
  const body = value(bodyKey);
  if (body) {
    if (body.length <= policy.maxMailTemplateLength && !containsHtml(body)) { result.mail.body = body; } else { ignored(bodyKey); }
  }
  return result;
}

export function splitLines(value: string): string[] {
  return value.split(/\r?\n/);
}

function createMailSettings(config: IFeedbackConfig, language: ILanguageSettings, uiLanguage: string, policy: ISecurityPolicy): IMailSettings {
  const mail = config.mail;
  const from = typeof mail.from === 'string' ? mail.from.trim() : '';
  const fromError = from ? checkEmail(from, policy) : '';
  if (fromError) {
    logError('Invalid sender address in the configuration, ignored', { setting: 'mail.from' });
  }
  return {
    subject: language.subject,
    body: language.body,
    cc: validEmails(Array.isArray(mail.cc) ? mail.cc : [], 'mail.cc', policy),
    bcc: validEmails(Array.isArray(mail.bcc) ? mail.bcc : [], 'mail.bcc', policy),
    from: fromError ? '' : from,
    replyToReporter: mail.replyToReporter !== false,
    locale: uiLanguage || language.code
  };
}

/** Exact match ("de-de"), then same language prefix ("de-at" -> "de-de"). */
function pickByLanguage<T>(entries: { [language: string]: T }, uiLanguage: string): T | undefined {
  const language = (uiLanguage || '').toLowerCase();
  if (entries[language]) {
    return entries[language];
  }
  const prefix = language.split('-')[0];
  const match = Object.keys(entries).filter((key: string) => key.split('-')[0] === prefix)[0];
  return match ? entries[match] : undefined;
}

/** Valid addresses of values read from SharePoint (separated by ";" or ","); the rest is logged. */
export function validEmails(values: string[], source: string, policy: ISecurityPolicy): string[] {
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  values
    .reduce((all: string[], value: string) => all.concat(value.split(/[;,]/)), [])
    .map((email: string) => email.trim())
    .filter((email: string) => !!email)
    .forEach((email: string) => {
      const key = email.toLowerCase();
      if (seen[key] || result.length >= policy.maxRecipientsPerField) {
        return;
      }
      const error = checkEmail(email, policy);
      if (error) {
        logWarning('Invalid address in SharePoint, ignored', { source, error });
        return;
      }
      seen[key] = true;
      result.push(email);
    });
  return result;
}

function safeFieldName(value: unknown, setting: string): string {
  const name = typeof value === 'string' ? value.trim() : '';
  if (!FIELD_NAME.test(name)) {
    logError('Invalid column name in the configuration', { setting, value: name });
    return '';
  }
  return name;
}

/** At most 255 characters (Title column). */
function safeEntryKey(value: unknown): string {
  const key = typeof value === 'string' ? singleLine(value) : '';
  return key.length <= 255 ? key : '';
}

/** undefined if no usable category, so that the defaults apply. */
function safeCategories(value: unknown, policy: ISecurityPolicy): string[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  (value as unknown[]).forEach((item: unknown) => {
    const text = typeof item === 'string' ? singleLine(item).substr(0, policy.maxCategoryLength) : '';
    const key = text.toLowerCase();
    if (text && !seen[key] && result.length < policy.maxCategories) {
      seen[key] = true;
      result.push(text);
    }
  });
  return result.length > 0 ? result : undefined;
}
