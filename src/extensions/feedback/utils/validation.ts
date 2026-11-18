import { ISecurityPolicy } from '../config/securityConfig';
import { formatText } from '../loc/template';
import { FeedbackErrors, FeedbackField, IFeedbackForm } from '../models/IFeedback';
import * as strings from 'FeedbackApplicationCustomizerStrings';

/** Pragmatic check: local part, @, domain with at least one dot. No spaces, brackets or separators. */
const EMAIL = /^[^\s@<>()[\]\\,;:"]+@([a-z0-9-]+\.)+[a-z]{2,}$/i;

/** Fields in the order of the form (left column, then right column). Used to focus the first error. */
export const FIELD_ORDER: FeedbackField[] = ['category', 'subject', 'url', 'description', 'contactEmail', 'mailboxEmail'];

/** Returns an error message or '' for a valid address. */
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

/** Splits an address field ("a@x.de; b@y.de") into trimmed, non-empty addresses. */
export function splitEmails(value: string): string[] {
  return value.split(/[;,]/).map((email: string) => email.trim()).filter((email: string) => !!email);
}

/** Checks every address of an address field. Returns the first error or ''. */
export function checkEmailList(value: string, policy: ISecurityPolicy): string {
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

/** Only absolute http(s) addresses with a host. */
export function checkUrl(url: string, policy: ISecurityPolicy): string {
  if (url.length > policy.maxUrlLength) {
    return formatText(strings.errorTooLong, policy.maxUrlLength);
  }
  return /^https?:\/\/[^\s/?#]+[^\s]*$/i.test(url) ? '' : strings.errorUrl;
}

/** Validates the whole form. Required: category, subject, URL, description and at least one recipient. */
export function validateForm(form: IFeedbackForm, categories: string[], policy: ISecurityPolicy): FeedbackErrors {
  const errors: FeedbackErrors = {};
  const subject = form.subject.trim();
  const url = form.url.trim();
  const description = form.description.trim();
  const contact = splitEmails(form.contactEmail).join(';');
  const mailbox = splitEmails(form.mailboxEmail).join(';');

  if (categories.indexOf(form.category) < 0) {
    errors.category = strings.errorCategory;
  }

  if (!subject) {
    errors.subject = strings.errorRequired;
  } else if (subject.length > policy.maxSubjectLength) {
    errors.subject = formatText(strings.errorTooLong, policy.maxSubjectLength);
  }

  errors.url = url ? checkUrl(url, policy) : strings.errorRequired;

  if (!description) {
    errors.description = strings.errorRequired;
  } else if (description.length > policy.maxDescriptionLength) {
    errors.description = formatText(strings.errorTooLong, policy.maxDescriptionLength);
  }

  if (!contact && !mailbox) {
    errors.contactEmail = strings.errorRecipient;
    errors.mailboxEmail = strings.errorRecipient;
  } else {
    errors.contactEmail = checkEmailList(contact, policy);
    errors.mailboxEmail = checkEmailList(mailbox, policy);
  }

  // Remove empty entries: a field without a key is valid.
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
