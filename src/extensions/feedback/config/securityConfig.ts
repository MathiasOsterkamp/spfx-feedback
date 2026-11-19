const SECURITY_SETTINGS = {
  allowedRecipientDomains: [] as string[],
  maxEmailLength: 254,
  maxRecipientsPerField: 10,
  maxTextLength: 255,
  maxDescriptionLength: 5000,
  maxCategories: 30,
  maxCategoryLength: 100,
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
