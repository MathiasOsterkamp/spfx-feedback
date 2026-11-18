/**
 * Security settings. Deliberately code-only (not in extension properties or the settings list):
 * site owners must not be able to relax them.
 */
const SECURITY_SETTINGS = {
  /**
   * Allowed recipient domains (incl. subdomains), e.g. ['firma.de']. Empty = all domains.
   * Recommended for production: visitors can edit the recipients, so without it the form can send
   * mails with SharePoint as sender to any address.
   */
  allowedRecipientDomains: [] as string[],
  /** RFC 5321 */
  maxEmailLength: 254,
  /** Addresses per field (contact, mailbox). */
  maxRecipientsPerField: 10,
  /** Single-line fields in form and settings view; 255 = SharePoint single line of text. */
  maxTextLength: 255,
  maxDescriptionLength: 5000,
  maxCategories: 30,
  maxCategoryLength: 100,
  /** Mail body template in the settings list. */
  maxMailTemplateLength: 20000
};

export interface ISecurityPolicy {
  allowedRecipientDomains: string[];
  maxEmailLength: number;
  maxRecipientsPerField: number;
  maxTextLength: number;
  maxDescriptionLength: number;
  maxCategories: number;
  maxCategoryLength: number;
  maxMailTemplateLength: number;
}

export function createSecurityPolicy(): ISecurityPolicy {
  return {
    allowedRecipientDomains: SECURITY_SETTINGS.allowedRecipientDomains
      .map((d: string) => d.trim().toLowerCase().replace(/^@/, ''))
      .filter((d: string) => !!d),
    maxEmailLength: SECURITY_SETTINGS.maxEmailLength,
    maxRecipientsPerField: SECURITY_SETTINGS.maxRecipientsPerField,
    maxTextLength: SECURITY_SETTINGS.maxTextLength,
    maxDescriptionLength: SECURITY_SETTINGS.maxDescriptionLength,
    maxCategories: SECURITY_SETTINGS.maxCategories,
    maxCategoryLength: SECURITY_SETTINGS.maxCategoryLength,
    maxMailTemplateLength: SECURITY_SETTINGS.maxMailTemplateLength
  };
}
