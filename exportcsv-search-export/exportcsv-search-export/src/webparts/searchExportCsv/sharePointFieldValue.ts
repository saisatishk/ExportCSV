/**
 * SharePoint often returns field values as `type;#value` (e.g. calculated columns: `string;#text`).
 * Lookup fields use `id;#display`. This returns the display/value part for CSV export.
 */
const SP_FIELD_TYPE_PREFIX = /^(string|number|integer|datetime|boolean|double|object|lookup|user|url|choice|counter|modstat|note|text)$/i;

export function stripSharePointSerializedFieldValue(raw: string): string {
  const s = (raw || '').trim();
  if (!s) return '';
  const i = s.indexOf(';#');
  if (i === -1) return s;
  const head = s.slice(0, i);
  const tail = s.slice(i + 2);
  if (SP_FIELD_TYPE_PREFIX.test(head)) {
    return tail;
  }
  if (/^\d+$/.test(head)) {
    return tail;
  }
  return s;
}
