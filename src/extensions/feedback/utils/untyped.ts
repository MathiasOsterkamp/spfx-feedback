/** Object of unknown origin (stored properties, REST responses). Every value must be checked before use. */
export interface IUntyped {
  [key: string]: unknown;
}

/** Returns the value as an object whose properties can be read safely. Non-objects become {}. */
export function asRecord(value: unknown): IUntyped {
  return value && typeof value === 'object' ? value as IUntyped : {};
}

/** Single entity of a SharePoint REST response (verbose: { d: {...} }, otherwise the object itself). */
export function odataEntity(json: unknown): IUntyped {
  const root = asRecord(json);
  return root.d && typeof root.d === 'object' ? asRecord(root.d) : root;
}

/** Item list of a SharePoint REST response (verbose: { d: { results: [] } }, otherwise { value: [] }). */
export function odataItems(json: unknown): IUntyped[] {
  const root = asRecord(json);
  const items: unknown = Array.isArray(root.value) ? root.value : asRecord(root.d).results;
  return Array.isArray(items) ? (items as unknown[]).map((item: unknown) => asRecord(item)) : [];
}

/**
 * Values of a multi-value field as an array: plain array (nometadata/minimal), { results: [] } (verbose),
 * a single value becomes [value], null/undefined become [].
 */
export function asArray(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value;
  }
  const results: unknown = asRecord(value).results;
  if (Array.isArray(results)) {
    return results;
  }
  return value === undefined || value === null ? [] : [value];
}
