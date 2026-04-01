import { stripSharePointSerializedFieldValue } from './sharePointFieldValue';

export type IPreparedSearchCell = { keyLower: string; value: string };

/** One pass per row: normalize keys/values once for fast column lookups. */
export function prepareSearchRowCells(
  cells: Array<{ Key?: unknown; Value?: unknown }>,
  normalizeToString: (value: unknown) => string
): IPreparedSearchCell[] {
  const out: IPreparedSearchCell[] = [];
  for (let i = 0; i < cells.length; i++) {
    const keyLower = normalizeToString(cells[i]?.Key).toLowerCase();
    const value = stripSharePointSerializedFieldValue(normalizeToString(cells[i]?.Value));
    out.push({ keyLower, value });
  }
  return out;
}

export function getPreparedCellValueForColumn(prepared: IPreparedSearchCell[], columnNameLower: string): string {
  for (let i = 0; i < prepared.length; i++) {
    if (prepared[i].keyLower.indexOf(columnNameLower) !== -1) {
      return prepared[i].value;
    }
  }
  return '';
}

export function getPreparedCellValueForCandidates(
  prepared: IPreparedSearchCell[],
  candidates: string[]
): string | undefined {
  const lowered: string[] = [];
  for (let c = 0; c < candidates.length; c++) {
    lowered.push(candidates[c].toLowerCase());
  }
  for (let i = 0; i < prepared.length; i++) {
    const key = prepared[i].keyLower;
    for (let j = 0; j < lowered.length; j++) {
      if (key.indexOf(lowered[j]) !== -1) {
        return prepared[i].value;
      }
    }
  }
  return undefined;
}
