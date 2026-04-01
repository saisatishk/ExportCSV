/**
 * SharePoint search often returns dates as ISO strings. CSV export uses US-style MM/dd/yyyy
 * (calendar date in UTC for parsed timestamps).
 */

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function isDateManagedPropertyName(name: string): boolean {
  const lower = name.trim().toLowerCase();
  if (lower.length >= 7 && lower.slice(-7) === 'owsdate') return true;
  if (lower.indexOf('datetime') >= 0) return true;
  if (/^refinabledate\d*$/i.test(lower.replace(/[^a-z0-9]/g, ''))) return true;
  const compact = lower.replace(/[^a-z0-9]/g, '');
  return /(created|modified|lastmodifiedtime|startdate|enddate)$/.test(compact);
}

/** ISO date or common date/time shapes from search (avoid short numeric-only strings). */
function looksLikeIsoOrSearchDateString(value: string): boolean {
  const t = value.trim();
  if (t.length < 10) return false;
  return /^\d{4}-\d{2}-\d{2}/.test(t);
}

function shouldFormatAsDate(managedPropertyName: string, raw: string): boolean {
  if (isDateManagedPropertyName(managedPropertyName)) return true;
  return looksLikeIsoOrSearchDateString(raw);
}

/** Format for CSV when column/value looks like a date; otherwise returns `raw` unchanged. */
export function formatCsvDateCell(managedPropertyName: string, raw: string): string {
  const s = (raw || '').trim();
  if (!s) return '';
  if (!shouldFormatAsDate(managedPropertyName, s)) return raw;
  const d = new Date(s);
  if (isNaN(d.getTime())) return raw;
  const mm = pad2(d.getUTCMonth() + 1);
  const dd = pad2(d.getUTCDate());
  const yyyy = String(d.getUTCFullYear());
  return `${mm}/${dd}/${yyyy}`;
}
