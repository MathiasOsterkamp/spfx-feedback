import { logWarning } from './logging';
import { singleLine } from './sanitize';
import { splitEmails } from './validation';
import { IFeedbackForm, IFeedbackOrigin, IMailSettings } from '../models/IFeedback';

/** Everything SP.Utilities.Utility.SendEmail can take. */
export interface IMail {
  to: string[];
  cc: string[];
  bcc: string[];
  /** Empty = SharePoint's default sender */
  from: string;
  /** Empty = no Reply-To header */
  replyTo: string;
  subject: string;
  /** HTML. All values from the form are escaped. */
  body: string;
}

/** Placeholders of the mail templates (config mail.templates, settings list). */
export type MailPlaceholder = 'category' | 'subject' | 'description' | 'url'
  | 'reporterName' | 'reporterEmail' | 'siteTitle' | 'siteUrl' | 'date' | 'recipients';

export const MAIL_PLACEHOLDERS: MailPlaceholder[] = [
  'category', 'subject', 'description', 'url', 'reporterName', 'reporterEmail', 'siteTitle', 'siteUrl', 'date', 'recipients'
];

const PLACEHOLDER = /\{\{\s*([A-Za-z]+)\s*\}\}/g;

/** Placeholders of a template that renderTemplate does not know (shown by the setup view). */
export function unknownPlaceholders(template: string): string[] {
  const unknown: string[] = [];
  template.replace(PLACEHOLDER, (match: string, name: string) => {
    if (MAIL_PLACEHOLDERS.indexOf(name as MailPlaceholder) < 0 && unknown.indexOf(match) < 0) {
      unknown.push(match);
    }
    return match;
  });
  return unknown;
}

/** SharePoint limits the subject; longer subjects are cut. */
const MAX_SUBJECT_LENGTH = 255;

/** Escapes text for HTML content and attribute values. */
export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Recipients without duplicates (case-insensitive), in the order contact(s), mailbox(es). */
export function recipients(form: IFeedbackForm): string[] {
  const seen: { [key: string]: boolean } = {};
  return splitEmails(form.contactEmail).concat(splitEmails(form.mailboxEmail))
    .filter((email: string) => {
      const key = email.toLowerCase();
      if (seen[key]) { return false; }
      seen[key] = true;
      return true;
    });
}

/**
 * Replaces {{placeholder}} in a template.
 * html = true: values are HTML-escaped, line breaks of the description become <br>.
 * Unknown placeholders are removed and logged (typo in the configuration).
 */
export function renderTemplate(template: string, values: { [K in MailPlaceholder]: string }, html: boolean): string {
  return template.replace(PLACEHOLDER, (match: string, name: string) => {
    if (!Object.prototype.hasOwnProperty.call(values, name)) {
      logWarning('Unknown placeholder in the mail template', { placeholder: match });
      return '';
    }
    const value = values[name as MailPlaceholder];
    if (!html) {
      return value;
    }
    const escaped = escapeHtml(value);
    return name === 'description' ? escaped.replace(/\r?\n/g, '<br>') : escaped;
  });
}

/** Date and time in the UI language; an unknown locale falls back to the browser default. */
function formatDate(date: Date, locale: string): string {
  try {
    return date.toLocaleString(locale);
  } catch {
    return date.toLocaleString();
  }
}

/**
 * Builds the notification mail from the configured template. The form must be validated before
 * (validateForm), so the URL is guaranteed to be http(s) and may be used in href.
 */
export function buildMail(form: IFeedbackForm, origin: IFeedbackOrigin, settings: IMailSettings, now: Date): IMail {
  const to = recipients(form);
  const values: { [K in MailPlaceholder]: string } = {
    category: form.category,
    subject: form.subject.trim(),
    description: form.description.trim(),
    url: form.url.trim(),
    reporterName: origin.userName,
    reporterEmail: origin.userEmail,
    siteTitle: origin.webTitle || origin.webUrl,
    siteUrl: origin.webUrl,
    date: formatDate(now, settings.locale),
    recipients: to.join('; ')
  };
  return {
    to,
    cc: settings.cc,
    bcc: settings.bcc,
    from: settings.from,
    replyTo: settings.replyToReporter ? origin.userEmail : '',
    // Plain text: no escaping, but no line breaks (header injection) and limited length.
    subject: singleLine(renderTemplate(settings.subject, values, false)).substr(0, MAX_SUBJECT_LENGTH),
    body: renderTemplate(settings.body, values, true)
  };
}
