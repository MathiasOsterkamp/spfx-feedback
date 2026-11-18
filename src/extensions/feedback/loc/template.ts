/** {0}, {1} ... in language file texts (plain data, they cannot contain functions). */
export function formatText(template: string, ...args: unknown[]): string {
  return template.replace(/\{(\d+)\}/g, (match: string, index: string) => {
    const value = args[parseInt(index, 10)];
    return value === undefined || value === null ? '' : String(value);
  });
}
