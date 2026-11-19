import { unknownPlaceholders } from './mail';
import { containsHtml, FIELD_NAME, LIST_NAME, singleLine, splitLines } from './sanitize';
import { ISecurityPolicy } from '../config/securityConfig';
import { formatText } from '../loc/template';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export function checkListName(value: string): string {
  const name = value.trim();
  return !name || LIST_NAME.test(name) ? '' : strings.setupErrorListName;
}

export function checkColumnName(value: string): string {
  const name = value.trim();
  return !name || FIELD_NAME.test(name) ? '' : strings.setupErrorColumnName;
}

export function checkCategories(value: string, policy: ISecurityPolicy): string {
  const lines = splitLines(value).map(singleLine).filter((line: string) => !!line);
  return lines.length > policy.maxCategories || lines.some((line: string) => line.length > policy.maxCategoryLength)
    ? formatText(strings.setupErrorCategories, policy.maxCategories, policy.maxCategoryLength)
    : '';
}

export function checkMailTemplate(value: string, maxLength: number): string {
  const text = value.trim();
  if (text.length > maxLength) {
    return formatText(strings.errorTooLong, maxLength);
  }
  if (containsHtml(text)) {
    return strings.setupErrorHtml;
  }
  const unknown = unknownPlaceholders(text);
  return unknown.length > 0 ? formatText(strings.setupErrorPlaceholder, unknown.join(', ')) : '';
}
