import { logWarning } from './logging';
import { singleLine } from './sanitize';
import { splitEmails } from './validation';
import { IFeedbackForm, IFeedbackOrigin, IMailSettings } from '../models/IFeedback';

export interface IMail {
  to: string[];
  cc: string[];
  bcc: string[];
  /** Empty = SharePoint's default sender */
  from: string;
  /** Empty = no Reply-To header */
  replyTo: string;
  subject: string;
  /** HTML; template text and all values are escaped. */
  body: string;
}

export type MailPlaceholder = 'category' | 'subject' | 'description' | 'url'
  | 'reporterName' | 'reporterEmail' | 'siteTitle' | 'siteUrl' | 'date' | 'recipients';

export const MAIL_PLACEHOLDERS: MailPlaceholder[] = [
  'category', 'subject', 'description', 'url', 'reporterName', 'reporterEmail', 'siteTitle', 'siteUrl', 'date', 'recipients'
];

const PLACEHOLDER = /\{\{\s*([A-Za-z]+)\s*\}\}/g;

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

/** SharePoint limits the subject. */
const MAX_SUBJECT_LENGTH = 255;

/** For HTML content and attribute values. */
function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Case-insensitive, without the ones in exclude. */
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

/** To: contact person(s); Cc: shared mailbox (contacts not repeated). */
export function recipients(form: IFeedbackForm): { to: string[]; cc: string[] } {
  const contacts = unique(splitEmails(form.contactEmail));
  return { to: contacts, cc: unique(splitEmails(form.mailboxEmail), contacts) };
}

/** Values are already escaped and validated (http(s) URL or e-mail address). */
const LINKS: { [K in MailPlaceholder]?: (escaped: string) => string } = {
  url: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  siteUrl: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  reporterEmail: (escaped: string) => (escaped ? `<a href="mailto:${escaped}">${escaped}</a>` : '')
};

/**
 * html = true: template text and values are HTML-escaped, so the template itself cannot inject HTML.
 * Unknown placeholders are removed and logged.
 */
function renderTemplate(template: string, values: { [K in MailPlaceholder]: string }, html: boolean): string {
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

/** Inline styles: mail programs ignore <style>. */
function htmlFrame(content: string): string {
  return `<div style="font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:#323130;">${content}</div>`;
}

/** An unknown locale throws: fall back to the browser default. */
function formatDate(date: Date, locale: string): string {
  try {
    return date.toLocaleString(locale);
  } catch {
    return date.toLocaleString();
  }
}

/** The form must be validated before (validateForm): the URL is then http(s) and safe in href. */
export function buildMail(form: IFeedbackForm, origin: IFeedbackOrigin, settings: IMailSettings, now: Date): IMail {
  const { to, cc: mailboxCc } = recipients(form);
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
    recipients: to.concat(mailboxCc).join('; ')
  };
  return {
    to,
    cc,
    bcc: settings.bcc,
    from: settings.from,
    replyTo: settings.replyToReporter ? origin.userEmail : '',
    // Plain text: no escaping, but no line breaks (header injection).
    subject: singleLine(renderTemplate(settings.subject, values, false)).substr(0, MAX_SUBJECT_LENGTH),
    body: htmlFrame(renderTemplate(settings.body, values, true))
  };
}
