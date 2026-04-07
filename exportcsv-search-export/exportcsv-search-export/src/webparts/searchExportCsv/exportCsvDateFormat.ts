/**
 * SharePoint search often returns dates as ISO strings. CSV export uses US-style MM/dd/yyyy
 * (calendar date in UTC for parsed timestamps).
 *
 * Which columns are treated as dates is controlled only by the web part property `csvDateColumns`
 * (comma-separated managed property names) — no name-pattern heuristics.
 */

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function shouldFormatAsDate(managedPropertyName: string, explicitDateColumnNames?: Set<string>): boolean {
  if (!explicitDateColumnNames || explicitDateColumnNames.size === 0) {
    return false;
  }
  const key = managedPropertyName.trim().toLowerCase();
  return explicitDateColumnNames.has(key);
}

/** Format for CSV when the column is listed in `csvDateColumns`; otherwise returns `raw` unchanged. */
export function formatCsvDateCell(
  managedPropertyName: string,
  raw: string,
  explicitDateColumnNames?: Set<string>
): string {
  const s = (raw || '').trim();
  if (!s) return '';
  if (!shouldFormatAsDate(managedPropertyName, explicitDateColumnNames)) return raw;
  const d = new Date(s);
  if (isNaN(d.getTime())) return raw;
  const mm = pad2(d.getUTCMonth() + 1);
  const dd = pad2(d.getUTCDate());
  const yyyy = String(d.getUTCFullYear());
  return `${mm}/${dd}/${yyyy}`;
}
