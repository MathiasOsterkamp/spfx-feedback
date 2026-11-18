import { logError, logWarning } from './logging';
import { IUntyped } from './untyped';
import { checkEmail } from './validation';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackSettings, IMailboxSettings, IMailSettings, ISettingsListSettings, IStoredValues } from '../models/IFeedback';

const COLOR_PATTERN = /^(#[0-9a-f]{3,8}|rgba?\(\s*[\d.\s,%]+\)|hsla?\(\s*[\d.\s,%deg]+\)|[a-z]{3,30})$/i;

/**
 * Internal field names consist of letters, digits and underscores (special characters are encoded
 * as _xHHHH_). The names are inserted into REST URLs, so nothing else is allowed.
 */
export const FIELD_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;

/** Name of a list in its URL ("Lists/<name>"); it becomes part of REST URLs. */
export const LIST_NAME = /^[A-Za-z0-9_-]{1,128}$/;

/** Allows only valid CSS color values (prevents CSS injection via color values). */
export function safeColor(value: unknown, fallback: string): string {
  const v = typeof value === 'string' ? value.trim() : '';
  return COLOR_PATTERN.test(v) ? v : fallback;
}

/** Removes control characters (incl. line breaks) and surrounding whitespace. */
export function singleLine(value: string): string {
  // eslint-disable-next-line no-control-regex -- control characters are exactly what is removed here
  return value.replace(/[\u0000-\u001f\u007f]+/g, ' ').trim();
}

/** Shape of config/feedbackConfig.ts (checked by createSettings). */
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
    storage: string;
    key: string;
  };
  mail: {
    cc: string[];
    bcc: string[];
    from: string;
    replyToReporter: boolean;
    templates: { [language: string]: { subject: string; body: string } };
  };
  defaultLanguage: string;
}

/** Used if the configuration contains no usable template (logged as error). */
const FALLBACK_TEMPLATE = { subject: '[Feedback] {{subject}}', body: '<p>{{category}}: {{subject}}</p><p>{{url}}</p><p>{{description}}</p><p>{{reporterName}} {{reporterEmail}}</p>' };

/** List URL relative to the site: "Lists/<name>" (name as in LIST_NAME). */
const LIST_URL = /^Lists\/[A-Za-z0-9_-]{1,128}$/;

/**
 * Checks the configuration from config/feedbackConfig.ts and picks the mail template of the UI language.
 * The categories come from the language files (strings.categories) and are only limited here.
 * Configuration mistakes (invalid names) are logged as errors; the affected feature is switched off
 * instead of sending broken requests. The values of the settings list are applied later (applyStoredValues).
 */
export function createSettings(config: IFeedbackConfig, uiLanguage: string, categories: string[], policy: ISecurityPolicy): IFeedbackSettings {
  const contact = config.contact;
  const columnName = safeFieldName(contact.pageColumn.internalName, 'contact.pageColumn.internalName');
  const emailName = safeFieldName(contact.emailColumn.internalName, 'contact.emailColumn.internalName');
  const listUrl = LIST_URL.test(contact.list.url) ? contact.list.url : '';
  if (!listUrl) {
    logError('Invalid configuration, contact disabled', { setting: 'contact.list.url', expected: 'Lists/<name>' });
  }
  const contactValid = !!columnName && !!emailName && !!listUrl;
  const settingsList = createSettingsList(config);
  return {
    categories: safeCategories(categories, policy) || ['-'],
    contactColumn: {
      internalName: contactValid ? columnName : '',
      displayName: singleLine(contact.pageColumn.displayName || '') || columnName
    },
    contactList: { url: listUrl, title: singleLine(contact.list.title || '') || listUrl.substr(6) },
    contactEmailColumn: { internalName: emailName, displayName: singleLine(contact.emailColumn.displayName || '') || emailName },
    settingsList,
    mailbox: createMailboxSettings(config, settingsList),
    mail: createMailSettings(config, uiLanguage, policy)
  };
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
      contactList: safePropertyKey(keys.contactList),
      contactColumn: safePropertyKey(keys.contactColumn),
      categories: safePropertyKey(keys.categories),
      mailSubject: safePropertyKey(keys.mailSubject),
      mailBody: safePropertyKey(keys.mailBody)
    }
  };
}

function createMailboxSettings(config: IFeedbackConfig, settingsList: ISettingsListSettings): IMailboxSettings {
  const storage = config.mailbox.storage === 'propertyBag' ? 'propertyBag' : 'list';
  const key = safePropertyKey(config.mailbox.key);
  // An invalid settings list is already logged: no mailbox from there.
  return { storage, key: storage === 'list' && !settingsList.url ? '' : key };
}

/**
 * Applies the entries of the settings list to the code configuration. Empty entries keep the code value.
 * Invalid entries are logged and ignored - the setup view does not store them, but the list can be edited directly.
 */
export function applyStoredValues(base: IFeedbackSettings, stored: IStoredValues, policy: ISecurityPolicy): IFeedbackSettings {
  const keys = base.settingsList.keys;
  const value = (key: string): string => (key && typeof stored[key] === 'string' ? stored[key].trim() : '');
  const ignored = (key: string): void => logError('Invalid value in the settings list, ignored', { key });
  const result: IFeedbackSettings = { ...base, mail: { ...base.mail } };

  const categories = value(keys.categories);
  if (categories) {
    const parsed = safeCategories(splitLines(categories), policy);
    if (parsed) { result.categories = parsed; } else { ignored(keys.categories); }
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

  // An empty internal name in the base means: the code configuration is invalid, the contact stays off.
  const columnName = value(keys.contactColumn);
  if (columnName && base.contactColumn.internalName) {
    if (!FIELD_NAME.test(columnName)) {
      ignored(keys.contactColumn);
    } else if (columnName !== base.contactColumn.internalName) {
      result.contactColumn = { internalName: columnName, displayName: columnName };
    }
  }

  const subject = value(keys.mailSubject);
  if (subject) {
    if (subject.length <= policy.maxSubjectLength) { result.mail.subject = singleLine(subject); } else { ignored(keys.mailSubject); }
  }
  const body = value(keys.mailBody);
  if (body) {
    if (body.length <= policy.maxMailTemplateLength) { result.mail.body = body; } else { ignored(keys.mailBody); }
  }
  return result;
}

/** Lines of a multi-line text value (categories: one per line). */
export function splitLines(value: string): string[] {
  return value.split(/\r?\n/);
}

function createMailSettings(config: IFeedbackConfig, uiLanguage: string, policy: ISecurityPolicy): IMailSettings {
  const mail = config.mail;
  const usable = (t: unknown): t is { subject: string; body: string } =>
    !!t && typeof (t as IUntyped).subject === 'string' && !!((t as IUntyped).subject as string).trim()
    && typeof (t as IUntyped).body === 'string' && !!((t as IUntyped).body as string).trim();
  const own = pickByLanguage(mail.templates, uiLanguage);
  const fallback = mail.templates[config.defaultLanguage];
  let template = usable(own) ? own : (usable(fallback) ? fallback : undefined);
  if (!template) {
    logError('No usable mail template in the configuration, using a minimal one', { setting: 'mail.templates' });
    template = FALLBACK_TEMPLATE;
  }
  const from = typeof mail.from === 'string' ? mail.from.trim() : '';
  const fromError = from ? checkEmail(from, policy) : '';
  if (fromError) {
    logError('Invalid sender address in the configuration, ignored', { setting: 'mail.from' });
  }
  return {
    subject: template.subject,
    body: template.body,
    cc: validEmails(Array.isArray(mail.cc) ? mail.cc : [], 'mail.cc', policy),
    bcc: validEmails(Array.isArray(mail.bcc) ? mail.bcc : [], 'mail.bcc', policy),
    from: fromError ? '' : from,
    replyToReporter: mail.replyToReporter !== false,
    locale: uiLanguage || config.defaultLanguage
  };
}

/** Entry for the UI language: exact match ("de-de"), then same language prefix ("de-at" -> "de-de"). */
function pickByLanguage<T>(entries: { [language: string]: T }, uiLanguage: string): T | undefined {
  const language = (uiLanguage || '').toLowerCase();
  if (entries[language]) {
    return entries[language];
  }
  const prefix = language.split('-')[0];
  const match = Object.keys(entries).filter((key: string) => key.split('-')[0] === prefix)[0];
  return match ? entries[match] : undefined;
}

/**
 * Keeps the valid addresses of a value read from SharePoint (field or property bag) and logs the rest.
 * A text value may contain several addresses separated by ";" or ",".
 */
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

/** Property bag keys are encoded for the request (see RecipientService), only the length is limited. */
function safePropertyKey(value: unknown): string {
  const key = typeof value === 'string' ? singleLine(value) : '';
  return key.length <= 255 ? key : '';
}

/** Returns undefined if the value contains no usable category, so that the defaults apply. */
export function safeCategories(value: unknown, policy: ISecurityPolicy): string[] | undefined {
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
