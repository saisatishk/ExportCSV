import { BaseWebComponent } from '@pnp/modern-search-extensibility';

interface IExportButtonAttributes {
  queryModification?: string;
  sourceId?: string;
}

interface IExportPageResult {
  rows: Array<{ Title: string; Path: string; Author: string }>;
  lastDocId?: number;
  totalRows?: number;
}

function escapeCsvCell(value: string): string {
  const v = value ?? '';
  const needsQuotes = /[",\r\n]/.test(v);
  const escaped = v.replace(/"/g, '""');
  return needsQuotes ? `"${escaped}"` : escaped;
}

function normalizeToString(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (typeof value === 'object') {
    const v = value as { value?: unknown; Value?: unknown; StringVal?: unknown };
    if (typeof v.value === 'string' || typeof v.value === 'number') return String(v.value);
    if (typeof v.Value === 'string' || typeof v.Value === 'number') return String(v.Value);
    if (typeof v.StringVal === 'string' || typeof v.StringVal === 'number') return String(v.StringVal);
  }
  return String(value);
}

type ICellLike = { Key?: unknown; Value?: unknown };

function getCellValue(cells: ICellLike[], keyCandidates: string[]): string | undefined {
  if (!cells || cells.length === 0) return undefined;

  const candidatesLower = [];
  for (let i = 0; i < keyCandidates.length; i++) {
    candidatesLower.push(String(keyCandidates[i]).toLowerCase());
  }

  for (let i = 0; i < cells.length; i++) {
    const c = cells[i];
    const cellKey = normalizeToString(c?.Key).toLowerCase();

    for (let j = 0; j < candidatesLower.length; j++) {
      // Use substring match to tolerate aliases/casing differences.
      if (cellKey.indexOf(candidatesLower[j]) !== -1) {
        return normalizeToString(c?.Value);
      }
    }
  }

  return undefined;
}

export class SearchExportCsvButtonWebComponent extends BaseWebComponent {
  private _isCancelled: boolean = false;

  public constructor() {
    super();
  }

  public async connectedCallback(): Promise<void> {
    // Clear whatever might exist (PnP recreates this component when attributes change).
    this.innerHTML = '';

    const container = document.createElement('div');

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = 'Export CSV';

    const cancelButton = document.createElement('button');
    cancelButton.type = 'button';
    cancelButton.textContent = 'Cancel';
    cancelButton.style.marginLeft = '8px';
    cancelButton.disabled = true;

    const status = document.createElement('div');
    status.style.marginTop = '8px';

    container.appendChild(button);
    container.appendChild(cancelButton);
    container.appendChild(status);
    this.appendChild(container);

    const showError = (message: string): void => {
      button.disabled = false;
      cancelButton.disabled = true;
      status.style.color = '#a4262c';
      status.textContent = message;
    };

    let queryModification = '';
    let sourceId = '';
    try {
      const attrs = this.resolveAttributes() as IExportButtonAttributes;

      let queryModificationRaw = (attrs.queryModification ?? '').trim();

      // If the Handlebars attribute was passed using JSONstringify, we may receive a JSON-encoded string.
      try {
        const maybeParsed = JSON.parse(queryModificationRaw);
        if (typeof maybeParsed === 'string') {
          queryModificationRaw = maybeParsed;
        }
      } catch {
        // ignore (it's already a plain string)
      }

      queryModification = queryModificationRaw;
      sourceId = (attrs.sourceId ?? '').trim();
    } catch (e) {
      showError(`Export CSV: attribute parsing failed: ${e instanceof Error ? e.message : String(e)}`);
      return;
    }

    if (!queryModification) {
      showError('Export CSV: missing queryModification for this web part instance.');
      return;
    }

    if (!sourceId) {
      showError('Export CSV: missing sourceId for this web part instance.');
      return;
    }

    button.onclick = async () => {
      this._isCancelled = false;
      button.disabled = true;
      cancelButton.disabled = false;
      status.style.color = '';
      status.textContent = 'Export started...';

      try {
        const pageInfo = (globalThis as unknown as {
          _spPageContextInfo?: { webAbsoluteUrl?: string };
        })._spPageContextInfo;
        const webUrl = (pageInfo?.webAbsoluteUrl ?? window.location.origin).replace(/\/$/, '');

        const columns = ['Title', 'Path', 'Author'];
        const docIdColumn = 'DocId';
        const selectProperties = [...columns, docIdColumn].join(',');

        const pageSize = 500; // SharePoint search paging limit per request
        const maxRows = 200000; // Safety cap to avoid runaway exports in-browser

        let exported = 0;
        let totalRows: number | undefined = undefined;
        let lastDocId: number | undefined = undefined;
        let pageIndex = 0;

        const csvLines: string[] = [];
        csvLines.push(columns.map(escapeCsvCell).join(','));

        let shouldContinue = true;
        while (shouldContinue) {
          if (this._isCancelled) {
            status.textContent = 'Export cancelled.';
            return;
          }

          pageIndex++;

          const queryText =
            lastDocId === undefined ? queryModification : `(${queryModification}) AND IndexDocId>${lastDocId}`;

          const pageResult = await this._fetchExportPage({
            webUrl,
            sourceId,
            queryText,
            pageSize,
            selectProperties
          });

          totalRows = totalRows ?? pageResult.totalRows;

          for (const r of pageResult.rows) {
            csvLines.push(
              [r.Title, r.Path, r.Author]
                .map((v) => escapeCsvCell(normalizeToString(v)))
                .join(',')
            );
          }

          exported += pageResult.rows.length;

          status.textContent = `Exporting... ${exported}${totalRows ? ` / ${totalRows}` : ''} (page ${pageIndex})`;

          if (!pageResult.lastDocId || pageResult.rows.length === 0) {
            shouldContinue = false;
            break;
          }
          if (exported >= maxRows) {
            status.textContent = `Export stopped at safety cap (${maxRows} rows).`;
            shouldContinue = false;
            break;
          }
          if (totalRows !== undefined && exported >= totalRows) {
            shouldContinue = false;
            break;
          }

          // Prepare next page
          lastDocId = pageResult.lastDocId;
        }

        const csvContent = `\uFEFF${csvLines.join('\r\n')}\r\n`;
        this._downloadFile('SearchResults.csv', csvContent);

        status.textContent = `Export complete. ${exported} rows exported.`;
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        showError(`Export CSV failed: ${message}`);
      } finally {
        button.disabled = false;
        cancelButton.disabled = true;
      }
    };

    cancelButton.onclick = () => {
      this._isCancelled = true;
      cancelButton.disabled = true;
      status.textContent = 'Cancelling...';
    };
  }

  private _downloadFile(fileName: string, csvContent: string): void {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';

    document.body.appendChild(a);
    a.click();
    a.remove();

    URL.revokeObjectURL(url);
  }

  private async _fetchExportPage(params: {
    webUrl: string;
    sourceId: string;
    queryText: string;
    pageSize: number;
    selectProperties: string;
  }): Promise<IExportPageResult> {
    const { webUrl, sourceId, queryText, pageSize, selectProperties } = params;

    const url =
      `${webUrl}/_api/search/query?` +
      `querytext=${encodeURIComponent(`'${queryText}'`)}` +
      `&sourceid=${encodeURIComponent(`'${sourceId}'`)}` +
      `&rowlimit=${pageSize}` +
      `&rowsperpage=${pageSize}` +
      `&selectproperties=${encodeURIComponent(`'${selectProperties}'`)}` +
      `&sortlist=${encodeURIComponent(`'[docid]:ascending'`)}`;

    const response = await fetch(url, {
      method: 'GET',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    });

    if (!response.ok) {
      throw new Error(`Search export request failed: HTTP ${response.status}`);
    }

    const json = await response.json();

    const d = json?.d;
    const primary =
      d?.postquery?.PrimaryQueryResult ||
      d?.query?.PrimaryQueryResult ||
      d?.postquery;

    const relevant = primary?.RelevantResults;
    const totalRows = relevant?.TotalRows;

    const tableRows =
      relevant?.Table?.Rows?.results ||
      relevant?.Table?.Rows;

    type IRowLike = { Cells?: { results?: ICellLike[] } | ICellLike[] };
    const rows: IRowLike[] = Array.isArray(tableRows) ? (tableRows as IRowLike[]) : [];

    const mapped = rows.map((row) => {
      const cells: ICellLike[] = Array.isArray(row?.Cells)
        ? (row.Cells as ICellLike[])
        : (row?.Cells?.results ?? []);

      return {
        Title: normalizeToString(getCellValue(cells, ['Title']) ?? ''),
        Path: normalizeToString(getCellValue(cells, ['Path']) ?? ''),
        Author: normalizeToString(getCellValue(cells, ['Author']) ?? '')
      };
    });

    // Get last DocId for next IndexDocId page.
    const lastRow = rows[rows.length - 1];
    if (!lastRow) {
      return { rows: mapped, totalRows };
    }
    const lastCells: ICellLike[] = Array.isArray(lastRow?.Cells)
      ? (lastRow?.Cells as ICellLike[])
      : ((lastRow?.Cells as { results?: ICellLike[] } | undefined)?.results ?? []);
    const lastDocIdRaw =
      getCellValue(lastCells, ['DocId', 'docid']);
    const lastDocIdNumber = lastDocIdRaw !== undefined && lastDocIdRaw !== '' ? Number(lastDocIdRaw) : undefined;
    const hasValidDocId =
      typeof lastDocIdNumber === 'number' && !isNaN(lastDocIdNumber);

    return {
      rows: mapped,
      lastDocId: hasValidDocId ? lastDocIdNumber : undefined,
      totalRows
    };
  }
}

export function ensureSearchExportCsvButtonRegistered(): boolean {
  const marker = '__search_export_csv_button_defined';
  const g = globalThis as unknown as { [key: string]: unknown };
  g[marker] = false;

  try {
    const ce = (globalThis as unknown as { customElements?: CustomElementRegistry }).customElements;
    if (!ce) return false;
    const tag = 'search-export-csv-button';
    if (ce.get(tag)) return true;

    ce.define(tag, SearchExportCsvButtonWebComponent);
    g[marker] = true;
    return true;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    g[marker] = `error:${msg}`;
    return false;
  }
}

