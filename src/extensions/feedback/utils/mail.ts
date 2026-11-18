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
  /** HTML, built from the plain-text template. Template text and all values are escaped. */
  body: string;
}

/** Placeholders of the mail templates (config mail.templates, settings list). */
export type MailPlaceholder = 'category' | 'subject' | 'description' | 'url'
  | 'reporterName' | 'reporterEmail' | 'siteTitle' | 'siteUrl' | 'date' | 'recipients';

export const MAIL_PLACEHOLDERS: MailPlaceholder[] = [
  'category', 'subject', 'description', 'url', 'reporterName', 'reporterEmail', 'siteTitle', 'siteUrl', 'date', 'recipients'
];

const PLACEHOLDER = /\{\{\s*([A-Za-z]+)\s*\}\}/g;

/** Placeholders of a template that renderTemplate does not know (shown by the settings view). */
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

/** Addresses without duplicates (case-insensitive) and without the ones in exclude. */
function unique(emails: string[], exclude: string[] = []): string[] {
  const seen: { [key: string]: boolean } = {};
  exclude.forEach((email: string) => { seen[email.toLowerCase()] = true; });
  return emails.filter((email: string) => {
    const key = email.toLowerCase();
    if (seen[key]) { return false; }
    seen[key] = true;
    return true;
  });
}

/**
 * Recipients of the mail: the contact person(s) in To (required, see validateForm), the shared mailbox in Cc
 * (optional; an address that is also a contact person is not repeated).
 */
export function recipients(form: IFeedbackForm): { to: string[]; cc: string[] } {
  const contacts = unique(splitEmails(form.contactEmail));
  return { to: contacts, cc: unique(splitEmails(form.mailboxEmail), contacts) };
}

/** Placeholders whose value becomes a link in the mail (values are http(s) URLs or an e-mail address). */
const LINKS: { [K in MailPlaceholder]?: (escaped: string) => string } = {
  url: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  siteUrl: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  reporterEmail: (escaped: string) => (escaped ? `<a href="mailto:${escaped}">${escaped}</a>` : '')
};

/**
 * Replaces {{placeholder}} in a plain-text template.
 * html = false: result is plain text (subject).
 * html = true: template text and values are HTML-escaped, line breaks become <br>, URLs and the
 * reporter's address become links. The template itself cannot contain HTML.
 * Unknown placeholders are removed and logged (typo in the configuration).
 */
export function renderTemplate(template: string, values: { [K in MailPlaceholder]: string }, html: boolean): string {
  const rendered = (html ? escapeHtml(template) : template).replace(PLACEHOLDER, (match: string, name: string) => {
    if (!Object.prototype.hasOwnProperty.call(values, name)) {
      logWarning('Unknown placeholder in the mail template', { placeholder: match });
      return '';
    }
    const value = values[name as MailPlaceholder];
    if (!html) {
      return value;
    }
    const escaped = escapeHtml(value);
    const link = LINKS[name as MailPlaceholder];
    return link ? link(escaped) : escaped;
  });
  return html ? rendered.replace(/\r?\n/g, '<br>') : rendered;
}

/** Frame of the mail body: font like the SharePoint UI (mail programs ignore <style>, so inline). */
function htmlFrame(content: string): string {
  return `<div style="font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:#323130;">${content}</div>`;
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
  const { to, cc: mailboxCc } = recipients(form);
  // Shared mailbox first, then the fixed copies of the configuration; nobody twice.
  const cc = unique(mailboxCc.concat(settings.cc), to);
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
    // The addresses of the form: contact person(s) and shared mailbox.
    recipients: to.concat(mailboxCc).join('; ')
  };
  return {
    to,
    cc,
    bcc: settings.bcc,
    from: settings.from,
    replyTo: settings.replyToReporter ? origin.userEmail : '',
    // Plain text: no escaping, but no line breaks (header injection) and limited length.
    subject: singleLine(renderTemplate(settings.subject, values, false)).substr(0, MAX_SUBJECT_LENGTH),
    body: htmlFrame(renderTemplate(settings.body, values, true))
  };
}
