export interface IUntyped {
  [key: string]: unknown;
}

export function asRecord(value: unknown): IUntyped {
  return value && typeof value === 'object' ? value as IUntyped : {};
}

export function odataEntity(json: unknown): IUntyped {
  const root = asRecord(json);
  return root.d && typeof root.d === 'object' ? asRecord(root.d) : root;
}

export function odataItems(json: unknown): IUntyped[] {
  const root = asRecord(json);
  const items: unknown = Array.isArray(root.value) ? root.value : asRecord(root.d).results;
  return Array.isArray(items) ? (items as unknown[]).map((item: unknown) => asRecord(item)) : [];
}

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
