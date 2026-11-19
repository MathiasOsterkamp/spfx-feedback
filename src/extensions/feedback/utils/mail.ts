import { logWarning } from './logging';
import { singleLine } from './sanitize';
import { splitEmails } from './validation';
import { IFeedbackForm, IFeedbackOrigin, IMailSettings } from '../models/IFeedback';

export interface IMail {
  to: string[];
  cc: string[];
  bcc: string[];
  from: string;
  replyTo: string;
  subject: string;
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

const MAX_SUBJECT_LENGTH = 255;

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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

export function recipients(form: IFeedbackForm): { to: string[]; cc: string[] } {
  const contacts = unique(splitEmails(form.contactEmail));
  return { to: contacts, cc: unique(splitEmails(form.mailboxEmail), contacts) };
}

const LINKS: { [K in MailPlaceholder]?: (escaped: string) => string } = {
  url: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  siteUrl: (escaped: string) => `<a href="${escaped}">${escaped}</a>`,
  reporterEmail: (escaped: string) => (escaped ? `<a href="mailto:${escaped}">${escaped}</a>` : '')
};

const EMPTY_BRACKETS = /[ \t]*\([ \t]*\)/g;

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
  }).replace(EMPTY_BRACKETS, '');
  return html ? rendered.replace(/\r?\n/g, '<br>') : rendered;
}

export function anonymousOrigin(origin: IFeedbackOrigin, name: string): IFeedbackOrigin {
  return { ...origin, userName: name, userEmail: '' };
}

function htmlFrame(content: string): string {
  return `<div style="font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:#323130;">${content}</div>`;
}

function formatDate(date: Date, locale: string): string {
  try {
    return date.toLocaleString(locale);
  } catch {
    return date.toLocaleString();
  }
}

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
    subject: singleLine(renderTemplate(settings.subject, values, false)).substr(0, MAX_SUBJECT_LENGTH),
    body: htmlFrame(renderTemplate(settings.body, values, true))
  };
}
