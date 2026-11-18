/**
 * CENTRAL SECURITY CONFIGURATION
 *
 * These values are deliberately NOT part of the extension properties or the settings list: site owners must not be able
 * to relax the rules. Changes are made here in code and rolled out with the package.
 */
export const SECURITY_SETTINGS = {
  /**
   * Allowed domains for the recipient addresses (incl. subdomains), e.g. ['firma.de'].
   * Empty = all domains allowed.
   *
   * The recipient fields are editable by every visitor. Without a restriction the form can be used
   * to send mails with SharePoint as sender to any address. Recommended for production.
   */
  allowedRecipientDomains: [] as string[],
  /** RFC 5321: an address has at most 254 characters. */
  maxEmailLength: 254,
  /** Addresses per field (contact, mailbox); the field is also limited to maxTextLength characters. */
  maxRecipientsPerField: 10,
  /**
   * Single-line text fields: subject, URL (only http(s) addresses), contact person and shared mailbox in the form;
   * single-line fields and the mail subject in the settings view. 255 = limit of a single line of text in SharePoint.
   */
  maxTextLength: 255,
  /** Description (multi-line) in the form. */
  maxDescriptionLength: 5000,
  /** Limits for the categories (language files and settings list). */
  maxCategories: 30,
  maxCategoryLength: 100,
  /** Mail body template stored in the settings list (plain text). */
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
