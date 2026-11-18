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
  /** Only http(s) addresses are accepted in the URL field. */
  maxUrlLength: 2048,
  /** RFC 5321: an address has at most 254 characters. */
  maxEmailLength: 254,
  /** Addresses per field (contact, mailbox). Person and lookup fields can hold several values. */
  maxRecipientsPerField: 10,
  maxSubjectLength: 255,
  maxDescriptionLength: 4000,
  /** Limits for the categories (language files and settings list). */
  maxCategories: 30,
  maxCategoryLength: 100,
  /** Mail body template stored in the settings list (HTML). */
  maxMailTemplateLength: 20000
};

export interface ISecurityPolicy {
  allowedRecipientDomains: string[];
  maxUrlLength: number;
  maxEmailLength: number;
  maxRecipientsPerField: number;
  maxSubjectLength: number;
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
    maxUrlLength: SECURITY_SETTINGS.maxUrlLength,
    maxEmailLength: SECURITY_SETTINGS.maxEmailLength,
    maxRecipientsPerField: SECURITY_SETTINGS.maxRecipientsPerField,
    maxSubjectLength: SECURITY_SETTINGS.maxSubjectLength,
    maxDescriptionLength: SECURITY_SETTINGS.maxDescriptionLength,
    maxCategories: SECURITY_SETTINGS.maxCategories,
    maxCategoryLength: SECURITY_SETTINGS.maxCategoryLength,
    maxMailTemplateLength: SECURITY_SETTINGS.maxMailTemplateLength
  };
}
