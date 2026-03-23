/** Parses comma-separated managed property names; dedupes case-insensitively; default Title,Path,Author. */
export function parseExportColumnKeys(raw: string | undefined): string[] {
  const fallback = ['Title', 'Path', 'Author'];
  const s = (raw || '').trim();
  if (!s) return fallback;
  const seen = new Set<string>();
  const out: string[] = [];
  const parts = s.split(',');
  for (let i = 0; i < parts.length; i++) {
    const name = parts[i].trim();
    if (!name) continue;
    const low = name.toLowerCase();
    if (seen.has(low)) continue;
    seen.add(low);
    out.push(name);
  }
  return out.length > 0 ? out : fallback;
}

/** Ensures IndexDocId/DocId is requested for paging; avoids duplicate names. */
export function mergeSelectPropertiesForExport(exportColumnKeys: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  const push = (name: string): void => {
    const low = name.toLowerCase();
    if (seen.has(low)) return;
    seen.add(low);
    out.push(name);
  };
  for (let i = 0; i < exportColumnKeys.length; i++) {
    push(exportColumnKeys[i]);
  }
  const hasPaging = exportColumnKeys.some(
    (k) => k.toLowerCase() === 'indexdocid' || k.toLowerCase() === 'docid'
  );
  if (!hasPaging) {
    push('IndexDocId');
  }
  return out;
}
