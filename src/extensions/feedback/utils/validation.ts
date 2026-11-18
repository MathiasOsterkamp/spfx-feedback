import { ISecurityPolicy } from '../config/securityConfig';
import { formatText } from '../loc/template';
import { FeedbackErrors, FeedbackField, IFeedbackForm } from '../models/IFeedback';
import * as strings from 'FeedbackApplicationCustomizerStrings';

/** Pragmatic check; no spaces, brackets or separators. */
const EMAIL = /^[^\s@<>()[\]\\,;:"]+@([a-z0-9-]+\.)+[a-z]{2,}$/i;

/** Order of the form; used to focus the first error. */
export const FIELD_ORDER: FeedbackField[] = ['category', 'subject', 'url', 'description', 'contactEmail', 'mailboxEmail'];

/** Returns an error message or ''. */
export function checkEmail(email: string, policy: ISecurityPolicy): string {
  if (email.length > policy.maxEmailLength || !EMAIL.test(email)) {
    return strings.errorEmail;
  }
  const domain = email.substr(email.lastIndexOf('@') + 1).toLowerCase();
  const allowed = policy.allowedRecipientDomains;
  if (allowed.length > 0 && !allowed.some((d: string) => domain === d || endsWith(domain, '.' + d))) {
    return strings.errorEmailDomain;
  }
  return '';
}

export function splitEmails(value: string): string[] {
  return value.split(/[;,]/).map((email: string) => email.trim()).filter((email: string) => !!email);
}

function checkEmailList(value: string, policy: ISecurityPolicy): string {
  const emails = splitEmails(value);
  if (emails.length > policy.maxRecipientsPerField) {
    return formatText(strings.errorTooManyEmails, policy.maxRecipientsPerField);
  }
  for (const email of emails) {
    const error = checkEmail(email, policy);
    if (error) {
      return error;
    }
  }
  return '';
}

/** Length is checked too: the pre-filled value from SharePoint can be longer than the field allows to type. */
export function checkAddressField(value: string, policy: ISecurityPolicy): string {
  return value.trim().length > policy.maxTextLength
    ? formatText(strings.errorTooLong, policy.maxTextLength)
    : checkEmailList(value, policy);
}

/** Only absolute http(s) addresses (the URL becomes a link in the mail). */
function checkUrl(url: string, policy: ISecurityPolicy): string {
  if (url.length > policy.maxTextLength) {
    return formatText(strings.errorTooLong, policy.maxTextLength);
  }
  return /^https?:\/\/[^\s/?#]+[^\s]*$/i.test(url) ? '' : strings.errorUrl;
}

export function validateForm(form: IFeedbackForm, categories: string[], policy: ISecurityPolicy): FeedbackErrors {
  const errors: FeedbackErrors = {};
  const subject = form.subject.trim();
  const url = form.url.trim();
  const description = form.description.trim();

  if (categories.indexOf(form.category) < 0) {
    errors.category = strings.errorCategory;
  }

  if (!subject) {
    errors.subject = strings.errorRequired;
  } else if (subject.length > policy.maxTextLength) {
    errors.subject = formatText(strings.errorTooLong, policy.maxTextLength);
  }

  errors.url = url ? checkUrl(url, policy) : strings.errorRequired;

  if (!description) {
    errors.description = strings.errorRequired;
  } else if (description.length > policy.maxDescriptionLength) {
    errors.description = formatText(strings.errorTooLong, policy.maxDescriptionLength);
  }

  // Contact person (To) is required, shared mailbox (Cc) optional.
  errors.contactEmail = splitEmails(form.contactEmail).length === 0
    ? strings.errorRecipient
    : checkAddressField(form.contactEmail, policy);
  errors.mailboxEmail = checkAddressField(form.mailboxEmail, policy);

  FIELD_ORDER.forEach((field: FeedbackField) => {
    if (!errors[field]) { delete errors[field]; }
  });
  return errors;
}

export function hasErrors(errors: FeedbackErrors): boolean {
  return FIELD_ORDER.some((field: FeedbackField) => !!errors[field]);
}

function endsWith(text: string, suffix: string): boolean {
  return text.length >= suffix.length && text.substr(text.length - suffix.length) === suffix;
}
