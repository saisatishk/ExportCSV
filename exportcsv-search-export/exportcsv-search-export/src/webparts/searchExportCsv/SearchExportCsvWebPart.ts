/* eslint-disable max-lines -- URL filter + search REST + export are intentionally in one web part */
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './SearchExportCsvWebPart.module.scss';
import * as strings from 'SearchExportCsvWebPartStrings';
import { mergeSelectPropertiesForExport, parseExportColumnKeys } from './exportColumnsConfig';
import {
  buildExportButtonStyleAttr,
  resolveButtonLabel,
  type IExportButtonAppearanceProps
} from './exportButtonAppearance';
import { buildExportButtonAppearanceGroupFields } from './exportButtonPropertyPane';
import { formatCsvDateCell } from './exportCsvDateFormat';
import {
  getPreparedCellValueForCandidates,
  getPreparedCellValueForColumn,
  prepareSearchRowCells
} from './searchExportCells';
/** Dispatched when `history.pushState` / `replaceState` run (PnP Modern Search updates filters this way). */
const SEARCH_EXPORT_LOCATION_CHANGE = 'searchExportCsvLocationChange';

/** Search REST returns at most ~500 rows per request; larger RowLimit only adds server work. */
const EXPORT_PAGE_SIZE = 500;

/** Safety cap for total exported rows. */
const MAX_EXPORT_ROWS = 200000;

/** One-time patch so URL-only updates re-render this web part. */
function ensureHistoryPatchForSearchExport(): void {
  const w = window as Window & { __searchExportCsvHistoryPatched?: boolean };
  if (w.__searchExportCsvHistoryPatched) {
    return;
  }
  w.__searchExportCsvHistoryPatched = true;
  const origPush = history.pushState.bind(history);
  const origReplace = history.replaceState.bind(history);
  history.pushState = (...args: Parameters<History['pushState']>): void => {
    origPush(...args);
    window.dispatchEvent(new Event(SEARCH_EXPORT_LOCATION_CHANGE));
  };
  history.replaceState = (...args: Parameters<History['replaceState']>): void => {
    origReplace(...args);
    window.dispatchEvent(new Event(SEARCH_EXPORT_LOCATION_CHANGE));
  };
}

export type ISearchExportCsvWebPartProps = {
  sourceId: string;
  exportColumns?: string;
  /** Comma-separated managed property names to always format as dates in CSV (optional). */
  csvDateColumns?: string;
  debugApi?: boolean;
} & IExportButtonAppearanceProps;

interface IPnpFilterValue {
  name?: string;
  value?: string;
  operator?: number;
}

interface IPnpFilterGroup {
  filterName?: string;
  values?: IPnpFilterValue[];
  operator?: string;
}

/** PnP `FilterComparisonOperator` mirror; URL `f` JSON stores the numeric enum per filter value. */
const enum PnpFilterComparisonOperator {
  Eq = 0,
  Neq = 1,
  Gt = 2,
  Lt = 3,
  Geq = 4,
  Leq = 5,
  Contains = 6
}

export default class SearchExportCsvWebPart extends BaseClientSideWebPart<ISearchExportCsvWebPartProps> {
  private _isCancelled: boolean = false;
  /** Tracks URL (search+hash) last shown in the UI so we can skip redundant re-renders. */
  private _lastUrlFingerprint: string = '';
  private _urlRefreshTimer: number | undefined;
  private _urlListenersBound: boolean = false;

  private readonly _onLocationMaybeChanged = (): void => {
    if (this._urlRefreshTimer !== undefined) {
      window.clearTimeout(this._urlRefreshTimer);
    }
    this._urlRefreshTimer = window.setTimeout(() => {
      this._urlRefreshTimer = undefined;
      const fp = `${window.location.search}|${window.location.hash}`;
      if (fp === this._lastUrlFingerprint) {
        return;
      }
      this.render();
    }, 100);
  };

  protected onInit(): Promise<void> {
    const sup = super.onInit();
    return (sup || Promise.resolve()).then(() => {
      ensureHistoryPatchForSearchExport();
      if (!this._urlListenersBound) {
        this._urlListenersBound = true;
        window.addEventListener(SEARCH_EXPORT_LOCATION_CHANGE, this._onLocationMaybeChanged);
        window.addEventListener('hashchange', this._onLocationMaybeChanged);
        window.addEventListener('popstate', this._onLocationMaybeChanged);
      }
    });
  }

  protected onDispose(): void {
    if (this._urlListenersBound) {
      this._urlListenersBound = false;
      window.removeEventListener(SEARCH_EXPORT_LOCATION_CHANGE, this._onLocationMaybeChanged);
      window.removeEventListener('hashchange', this._onLocationMaybeChanged);
      window.removeEventListener('popstate', this._onLocationMaybeChanged);
    }
    if (this._urlRefreshTimer !== undefined) {
      window.clearTimeout(this._urlRefreshTimer);
      this._urlRefreshTimer = undefined;
    }
    super.onDispose();
  }

  private _formatUnknownForDebug(value: unknown): string {
    if (value === null) return 'null';
    if (value === undefined) return 'undefined';
    if (typeof value === 'string') return value.length > 200 ? `${value.slice(0, 200)}…` : value;
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  /** Property pane `csvDateColumns` — comma-separated managed property names for CSV date formatting. */
  private _getCsvExplicitDateColumns(): Set<string> | undefined {
    const raw = (this.properties.csvDateColumns || '').trim();
    if (!raw) {
      return undefined;
    }
    const set = new Set<string>();
    const parts = raw.split(',');
    for (let i = 0; i < parts.length; i++) {
      const t = parts[i].trim().toLowerCase();
      if (t) {
        set.add(t);
      }
    }
    return set.size > 0 ? set : undefined;
  }

  /** Logs to browser DevTools (F12); pass `Error` instances for stack traces. */
  private _logSearchExportError(context: string, error: unknown): void {
    if (error instanceof Error) {
      console.error(`[SearchExportCsv] ${context}: ${error.message}`, error);
    } else {
      console.error(`[SearchExportCsv] ${context}:`, error);
    }
  }

  /** Zero rows is not an exception — use this so users still see DevTools hints. */
  private _logSearchExportWarn(message: string, details?: Record<string, unknown>): void {
    if (details) {
      console.warn(`[SearchExportCsv] ${message}`, details);
    } else {
      console.warn(`[SearchExportCsv] ${message}`);
    }
  }

  /**
   * When the first export page has no rows, logs structured diagnostics (not only on thrown errors).
   * Turn on the web part property "debug API" for the same details in the status line.
   */
  private _logSearchExportZeroRowsFirstPage(
    debug: {
      sentQueryText: string;
      sentRefinementFilters: string;
      sentSourceId: string;
      extractedRows: number;
      totalRowsRawType: string;
      totalRowsRawValue: string;
      tableRowsIsArray: boolean;
      tableRowsHasResultsArray: boolean;
      tableRowsResultsLength?: number;
      primaryPath: string;
      relevantDefined: boolean;
      relevantHow: string;
      odataAttempt: string;
      jsonTopKeys: string;
      transport: string;
    },
    ctx: {
      effectiveQuery: string;
      refinementFilters?: string[];
      sourceId: string;
      pageUrl: string;
    }
  ): void {
    const trParsed = parseInt(String(debug.totalRowsRawValue).trim(), 10);
    const apiReportsHits = !isNaN(trParsed) && trParsed > 0;
    let hint: string;
    if (apiReportsHits && debug.extractedRows === 0) {
      hint =
        'Search returned TotalRows > 0 but no table rows were mapped — response shape may differ on this tenant; check jsonKeys / primaryPath.';
    } else if ((debug.sentRefinementFilters || '').trim()) {
      hint =
        'With refiners: confirm the Result Source ID matches the PnP Search Results web part. Mismatch often yields 0 rows here while the other web part still shows hits.';
    } else {
      hint = 'Confirm query text and Result Source ID match the Search Results web part.';
    }
    this._logSearchExportWarn(`Export first page returned 0 rows. ${hint}`, {
      debug,
      effectiveQuery: ctx.effectiveQuery,
      refinementFilters: ctx.refinementFilters,
      sourceId: ctx.sourceId,
      pageUrl: ctx.pageUrl
    });
  }

  public render(): void {
    try {
    // Keywords: URL (`k`, `q`, …) → discovery → page Search box → `*` at export time (no property-pane KQL).
    const effectiveQueryText = this._resolveSearchQueryForExport();
    const effectiveSourceId = this._resolveValue(this.properties.sourceId, undefined, 'sourceid');
    const filterParts = this._getUrlFilterParts();
    const combinedFilterHint = (() => {
      const parts: string[] = [];
      if (filterParts.refinementFql) {
        parts.push(`FQL Refinement: ${filterParts.refinementFql}`);
      }
      if (filterParts.filterKql) {
        parts.push(filterParts.filterKql);
      }
      return parts.length ? parts.map((p) => `(${p})`).join(' AND ') : '';
    })();

    const hasActiveRefinersForExport =
      !!((filterParts.refinementFql || '').trim()) || !!((filterParts.filterKql || '').trim());

    const noKeywordsHintMessage = !effectiveQueryText.value.trim()
      ? hasActiveRefinersForExport
        ? strings.ExportNoKeywordsWithRefinersHint
        : strings.ExportNoKeywordsNoRefinersHint
      : '';

    const showDebugUi = this.properties.debugApi === true;
    const sectionClass = `${styles.searchExportCsv}${showDebugUi ? '' : ` ${styles.searchExportCsvMinimal}`}`;
    const exportBtnLabel = resolveButtonLabel(this.properties.exportButtonText, strings.ExportButtonLabel);
    const cancelBtnLabel = resolveButtonLabel(this.properties.cancelButtonText, strings.CancelButtonLabel);
    const exportBtnStyle = buildExportButtonStyleAttr(this.properties);

    this.domElement.innerHTML = `
      <section class="${sectionClass}">
        ${
          showDebugUi
            ? `
        <div class="${styles.title}">${strings.WebPartTitle}</div>
        <div class="${styles.description}">${strings.WebPartDescription}</div>

        <div class="${styles.config}">
          <div><strong>${strings.ResolvedQueryLabel}:</strong> ${this._escapeHtml(effectiveQueryText.value || '*')}</div>
          <div><strong>${strings.SourceIdLabel}:</strong> ${this._escapeHtml(effectiveSourceId.value)}</div>
          <div class="${styles.hint}">
            ${strings.QuerySourceLabel}: ${effectiveQueryText.origin} | ${strings.SourceIdSourceLabel}: ${effectiveSourceId.origin}
          </div>
          ${
            noKeywordsHintMessage
              ? `<div class="${hasActiveRefinersForExport ? styles.hintNote : styles.hintWarning}">${this._escapeHtml(noKeywordsHintMessage)}</div>`
              : ''
          }
          <div class="${styles.hint}">
            ${strings.FiltersFromUrlLabel}: ${this._escapeHtml(filterParts.summary)}
          </div>
          ${
            combinedFilterHint
              ? `<div class="${styles.hint}"><strong>${strings.EffectiveFilterKqlLabel}:</strong> ${this._escapeHtml(combinedFilterHint)}</div>`
              : ''
          }
        </div>`
            : ''
        }

        <div class="${styles.actions}">
          <button type="button" class="${styles.button}" data-action="export"${exportBtnStyle}>${this._escapeHtml(exportBtnLabel)}</button>
          ${
            showDebugUi
              ? `<button type="button" class="${styles.button}" data-action="cancel" disabled${exportBtnStyle}>${this._escapeHtml(cancelBtnLabel)}</button>`
              : ''
          }
        </div>

        <div class="${styles.status}" data-role="status"></div>
      </section>
    `;

    const exportButton = this.domElement.querySelector<HTMLButtonElement>('button[data-action="export"]');
    const cancelButton = showDebugUi
      ? this.domElement.querySelector<HTMLButtonElement>('button[data-action="cancel"]')
      : null;
    const status = this.domElement.querySelector<HTMLDivElement>('div[data-role="status"]');

    if (!exportButton || !status) return;

    const showError = (message: string): void => {
      status.textContent = message;
      status.className = `${styles.status} ${styles.error}`;
      exportButton.disabled = false;
      if (cancelButton) {
        cancelButton.disabled = true;
      }
    };

    exportButton.onclick = async (): Promise<void> => {
      // Always read the latest URL + properties at click time. PnP updates `f` via history.replaceState
      // without reloading the page, so values captured during the last render() would stay stale otherwise.
      const liveQuery = this._resolveSearchQueryForExport();
      const liveSource = this._resolveValue(this.properties.sourceId, undefined, 'sourceid');
      const liveFilterParts = this._getUrlFilterParts();

      let queryBase = liveQuery.value.trim();
      const sourceId = liveSource.value.trim();
      const filterKql = liveFilterParts.filterKql;
      const refinementFql = (liveFilterParts.refinementFql || '').trim();
      const refinementFiltersPayload = refinementFql ? [refinementFql] : undefined;

      if (filterKql) {
        queryBase = queryBase ? `(${queryBase}) AND (${filterKql})` : filterKql;
      }

      // If the user removed the query text, export all from the configured SourceId.
      if (!queryBase) {
        // Use empty querytext (rather than `*`) for "match all".
        // Also note: we must handle pagination separately to avoid generating invalid `()` expressions.
        queryBase = '*';
      }
      if (!sourceId) {
        showError(strings.SourceIdRequiredError);
        return;
      }

      this._isCancelled = false;
      exportButton.disabled = true;
      if (cancelButton) {
        cancelButton.disabled = false;
      }
      status.textContent = strings.ExportStarted;
      status.className = `${styles.status}`;

      try {
        const columns = parseExportColumnKeys(this.properties.exportColumns);
        const selectPropertiesList = mergeSelectPropertiesForExport(columns);
        const selectProperties = selectPropertiesList.join(',');
        const pageSize = EXPORT_PAGE_SIZE;
        const maxRows = MAX_EXPORT_ROWS;
        let pageIndex = 0;
        let exported = 0;
        let totalRows: number | undefined;
        let lastDocId: number | undefined;
        let shouldContinue = true;
        let lastDebug: {
          sentQueryText: string;
          sentRefinementFilters: string;
          sentSourceId: string;
          extractedRows: number;
          totalRowsRawType: string;
          totalRowsRawValue: string;
          tableRowsIsArray: boolean;
          tableRowsHasResultsArray: boolean;
          tableRowsResultsLength?: number;
          primaryPath: string;
          relevantDefined: boolean;
          relevantHow: string;
          odataAttempt: string;
          jsonTopKeys: string;
          transport: string;
        } | undefined;

        const csvLines: string[] = [];
        csvLines.push(columns.join(','));
        const queryBaseTrim = queryBase.trim();
        const hasBaseQuery = queryBaseTrim.length > 0;
        let lastStatusAt = 0;

        while (shouldContinue) {
          if (this._isCancelled) {
            status.textContent = strings.ExportCancelled;
            return;
          }

          pageIndex++;
          const effectiveQuery =
            lastDocId === undefined
              ? queryBase
              : hasBaseQuery
                ? (queryBaseTrim === '*' ? `IndexDocId>${lastDocId}` : `(${queryBase}) AND IndexDocId>${lastDocId}`)
                : `IndexDocId>${lastDocId}`;

          const result = await this._fetchExportPage({
            webUrl: this.context.pageContext.web.absoluteUrl,
            sourceId,
            queryText: effectiveQuery,
            pageSize,
            selectProperties,
            selectPropertiesList,
            exportColumnKeys: columns,
            refinementFilters: refinementFiltersPayload,
            enableGetFallbackWhenEmpty: lastDocId === undefined
          });

          if (totalRows === undefined && result.totalRows !== undefined) {
            totalRows = result.totalRows;
          }
          if (this.properties.debugApi === true) {
            lastDebug = result.debug;
          }
          if (pageIndex === 1 && result.rows.length === 0) {
            this._logSearchExportZeroRowsFirstPage(result.debug, {
              effectiveQuery,
              refinementFilters: refinementFiltersPayload,
              sourceId,
              pageUrl: typeof window !== 'undefined' ? window.location.href : ''
            });
          }

          for (let i = 0; i < result.rows.length; i++) {
            const row = result.rows[i];
            const cells: string[] = [];
            for (let c = 0; c < columns.length; c++) {
              const key = columns[c];
              cells.push(this._escapeCsvCell(row[key] || ''));
            }
            csvLines.push(cells.join(','));
          }

          exported += result.rows.length;
          const nowMs = Date.now();
          if (nowMs - lastStatusAt >= 400 || pageIndex === 1) {
            lastStatusAt = nowMs;
            status.textContent = `${strings.ExportInProgress} ${exported}${totalRows ? ` / ${totalRows}` : ''} (${strings.PageLabel} ${pageIndex})`;
          }

          if (!result.lastDocId || result.rows.length === 0) {
            shouldContinue = false;
          } else if (exported >= maxRows) {
            status.textContent = `${strings.ExportCappedMessage} (${maxRows}).`;
            shouldContinue = false;
          } else if (totalRows !== undefined && exported >= totalRows) {
            shouldContinue = false;
          } else {
            lastDocId = result.lastDocId;
          }
        }

        this._downloadCsv('SearchResults.csv', `\uFEFF${csvLines.join('\r\n')}\r\n`);
        if (this.properties.debugApi && lastDebug) {
          status.textContent =
            `${strings.ExportCompleted} ${exported}` +
            `${totalRows !== undefined ? ` / ${totalRows}` : ''} ` +
            `(debug: sentQuery="${lastDebug.sentQueryText}", sentRefinement="${lastDebug.sentRefinementFilters}", sentSourceId=${lastDebug.sentSourceId}, ` +
            `extractedRows=${lastDebug.extractedRows}, ` +
            `tableRowsIsArray=${lastDebug.tableRowsIsArray}, tableRowsHasResultsArray=${lastDebug.tableRowsHasResultsArray}, tableRowsResultsLength=${lastDebug.tableRowsResultsLength ?? 'n/a'}, ` +
            `transport=${lastDebug.transport}, relevantHow=${lastDebug.relevantHow}, odata=${lastDebug.odataAttempt}, jsonKeys=${lastDebug.jsonTopKeys}, ` +
            `primaryPath=${lastDebug.primaryPath ?? 'n/a'}, relevantDefined=${lastDebug.relevantDefined}, ` +
            `totalRowsRaw=${lastDebug.totalRowsRawType}:${lastDebug.totalRowsRawValue}`;
        } else {
          status.textContent = `${strings.ExportCompleted} ${exported}${totalRows !== undefined ? ` / ${totalRows}` : ''}`;
        }
      } catch (error) {
        this._logSearchExportError('Export failed', error);
        const message = error instanceof Error ? error.message : String(error);
        showError(`${strings.ExportFailedPrefix} ${message}`);
      } finally {
        exportButton.disabled = false;
        if (cancelButton) {
          cancelButton.disabled = true;
        }
      }
    };

    if (cancelButton) {
      cancelButton.onclick = (): void => {
        this._isCancelled = true;
        cancelButton.disabled = true;
        status.textContent = strings.CancellingMessage;
        status.className = `${styles.status}`;
      };
    }
    } finally {
      this._lastUrlFingerprint = `${window.location.search}|${window.location.hash}`;
    }
  }

  /**
   * Read a query-string parameter from `?` and (when present) from the hash fragment
   * (`#k=...&f=...` is used on some modern search experiences).
   */
  private _getUrlParam(name: string): string {
    const key = (name || '').trim();
    if (!key) return '';

    const pageUrl = new URL(window.location.href);
    const fromSearch = (pageUrl.searchParams.get(key) || '').trim();
    if (fromSearch) return fromSearch;

    const rawHash = (window.location.hash || '').replace(/^#/, '');
    if (!rawHash || rawHash.indexOf('=') === -1) return '';

    try {
      const qp = rawHash.indexOf('?') >= 0 ? rawHash.split('?').slice(1).join('?') : rawHash;
      const hp = new URLSearchParams(qp);
      return (hp.get(key) || '').trim();
    } catch {
      return '';
    }
  }

  private _escapeHtml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private _resolveValue(propertyValue: string | undefined, paramName: string | undefined, fallbackParam: string): {
    value: string;
    origin: string;
  } {
    const direct = (propertyValue || '').trim();
    if (direct) {
      return { value: direct, origin: strings.FromPropertyLabel };
    }

    const key = (paramName || '').trim() || fallbackParam;
    const fromUrl = this._getUrlParam(key);
    if (fromUrl) {
      return { value: fromUrl, origin: `${strings.FromUrlLabel} (${key})` };
    }

    // Some pages use `k=` for the query text, while others still use `q=`.
    // If we already tried `k` (common in Modern Search), fall back to `q` as well.
    if (key === 'k') {
      const fromUrlQ = this._getUrlParam('q');
      if (fromUrlQ) {
        return { value: fromUrlQ, origin: `${strings.FromUrlLabel} (q)` };
      }
    }

    return { value: '', origin: strings.NotSetLabel };
  }

  /** Ordered URL keys for keyword query (PnP uses `k` when “Update URL” is enabled). Fixed list — no property-pane overrides. */
  private _searchQueryUrlParamCandidates(): string[] {
    return [
      'k',
      'q',
      'query',
      'search',
      'qt',
      'queryText',
      'keywords',
      'keyword',
      'sq',
      'text',
      'SearchQuery',
      'SearchText'
    ];
  }

  /**
   * Some tenants / web parts use an uncommon query param name. Scan ? and hash for likely keyword params (excluding `f` JSON).
   */
  private _tryDiscoverSearchQueryFromUrlParams(): { value: string; key: string } | undefined {
    const filtersKey = 'f';
    const skipExact = new Set(
      [filtersKey, 'f', 'sourceid', 'p', 'id', 'debug', 'env', 'locale'].map((x) => x.toLowerCase())
    );

    const entries = this._collectAllUrlParamEntries();
    const nameLooksLikeQuery = (key: string): boolean => {
      const kl = key.toLowerCase();
      if (/(^|_|-)(k|q)$/.test(kl)) return true;
      return (
        /search|query|keyword|sq$|st$|^text$|^terms$/i.test(kl) && !/filter|refiner|facet|source/i.test(kl)
      );
    };

    for (let i = 0; i < entries.length; i++) {
      const { key, value } = entries[i];
      const kl = key.trim().toLowerCase();
      if (!kl || skipExact.has(kl)) continue;

      const v = (value || '').trim();
      if (!v || v.length > 1500) continue;
      if (v.charAt(0) === '[' || v.charAt(0) === '{') continue;
      if (/^\{?[0-9a-f-]{36}\}?$/i.test(v)) continue;
      if (/\bfilterName\b/i.test(v)) continue;

      if (nameLooksLikeQuery(key)) {
        return { value: v, key };
      }
    }
    return undefined;
  }

  private _inputIsProbablyRefinerOrDialog(el: Element): boolean {
    if (el.closest('.ms-Panel-main, .ms-Dialog-main, [role="dialog"], [aria-modal="true"]')) {
      return true;
    }
    if (
      el.closest(
        '[class*="Refinement"], [class*="refinement"], [class*="filterPane"], [data-sp-fre-filter], [class*="ms-DatePicker"]'
      )
    ) {
      return true;
    }
    const ph = (el as HTMLInputElement).placeholder || '';
    if (/refiner|filter|from date|to date|start date|end date/i.test(ph)) {
      return true;
    }
    return false;
  }

  private _scoreSearchInputCandidate(el: HTMLInputElement | HTMLTextAreaElement): number {
    if (this._inputIsProbablyRefinerOrDialog(el)) {
      return 0;
    }
    const r = el.getBoundingClientRect();
    if (r.width < 48 || r.height < 16) {
      return 0;
    }
    if (r.bottom < 0 || r.top > (window.innerHeight || 800) + 200) {
      return 0;
    }
    const v = (el.value || '').trim();
    if (!v) {
      return 0;
    }
    let score = Math.min(v.length, 400);
    if (el instanceof HTMLInputElement && el.classList.contains('ms-SearchBox-field')) {
      score += 800;
    }
    if (el.closest('.ms-SearchBox')) {
      score += 400;
    }
    const ph = (el.placeholder || '').toLowerCase();
    if (ph.indexOf('search') !== -1) {
      score += 120;
    }
    const al = ((el as HTMLInputElement).getAttribute('aria-label') || '').toLowerCase();
    if (al.indexOf('search') !== -1) {
      score += 120;
    }
    return score;
  }

  /** Prefer page canvas / main so DOM walks skip suite chrome (faster on every render). */
  private _getPageContentRoots(): Element[] {
    const roots: Element[] = [];
    const canvas = document.querySelector('#spPageCanvasContent');
    const pageContent = document.querySelector('[data-sp-placeholder="PageContent"]');
    const main = document.querySelector('[role="main"]');
    if (canvas) {
      roots.push(canvas);
    }
    if (pageContent) {
      roots.push(pageContent);
    }
    if (main) {
      roots.push(main);
    }
    roots.push(document.body);
    return roots;
  }

  /**
   * Read keyword text from SharePoint / Fluent / PnP SearchBox-style fields (best-scoring visible input).
   */
  private _tryReadSharePointSearchBoxValue(): string {
    const roots = this._getPageContentRoots();

    const selector =
      [
        'input.ms-SearchBox-field',
        '.ms-SearchBox input[type="text"]',
        '.ms-SearchBox input[type="search"]',
        '.ms-SearchBox input',
        '[class*="SearchBox"] input[type="text"]',
        '[class*="SearchBox"] input[type="search"]',
        '[class*="searchBox"] input',
        'input[type="search"]',
        'input[placeholder*="Search" i]',
        'input[aria-label*="Search" i]',
        'textarea[placeholder*="Search" i]'
      ].join(', ');

    let bestVal = '';
    let bestScore = 0;

    for (let r = 0; r < roots.length; r++) {
      const fields = roots[r].querySelectorAll<HTMLInputElement | HTMLTextAreaElement>(selector);
      for (let i = 0; i < fields.length; i++) {
        const sc = this._scoreSearchInputCandidate(fields[i]);
        if (sc > bestScore) {
          bestScore = sc;
          bestVal = (fields[i].value || '').trim();
        }
      }
    }
    return bestVal;
  }

  /** Search keywords: URL (`k`, `q`, …) → loose URL scan → visible SearchBox. Empty → export uses `*`. */
  private _resolveSearchQueryForExport(): { value: string; origin: string } {
    const keys = this._searchQueryUrlParamCandidates();
    for (let k = 0; k < keys.length; k++) {
      const fromUrl = this._getUrlParam(keys[k]);
      if (fromUrl) {
        return { value: fromUrl, origin: `${strings.FromUrlLabel} (${keys[k]})` };
      }
    }

    const discovered = this._tryDiscoverSearchQueryFromUrlParams();
    if (discovered) {
      return {
        value: discovered.value,
        origin: `${strings.FromUrlLabel} (${discovered.key})`
      };
    }

    const fromBox = this._tryReadSharePointSearchBoxValue();
    if (fromBox) {
      return { value: fromBox, origin: strings.SearchQueryFromPageSearchBoxLabel };
    }

    return { value: '', origin: strings.NotSetLabel };
  }

  /** Collect `?` and hash query pairs (hash uses the same rules as `_getUrlParam`). */
  private _collectAllUrlParamEntries(): Array<{ key: string; value: string }> {
    const out: Array<{ key: string; value: string }> = [];
    const pageUrl = new URL(window.location.href);
    pageUrl.searchParams.forEach((value, key) => {
      const v = (value || '').trim();
      if (v) {
        out.push({ key, value: v });
      }
    });

    const rawHash = (window.location.hash || '').replace(/^#/, '');
    if (rawHash && rawHash.indexOf('=') !== -1) {
      try {
        const qp = rawHash.indexOf('?') >= 0 ? rawHash.split('?').slice(1).join('?') : rawHash;
        const hp = new URLSearchParams(qp);
        hp.forEach((value, key) => {
          const v = (value || '').trim();
          if (v) {
            out.push({ key, value: v });
          }
        });
      } catch {
        // ignore malformed hash query
      }
    }

    return out;
  }

  /** PnP sometimes double-encodes `f`; decode until stable or max passes. */
  private _decodePnpRefinerChain(raw: string): string {
    let s = raw || '';
    for (let pass = 0; pass < 4; pass++) {
      try {
        const next = decodeURIComponent(s);
        if (next === s) {
          break;
        }
        s = next;
      } catch {
        break;
      }
    }
    return s;
  }

  private _isPnpRefinerFilterJson(jsonText: string): boolean {
    const t = (jsonText || '').trim();
    if (!t || t.charAt(0) !== '[') {
      return false;
    }
    try {
      const parsed = JSON.parse(t) as unknown;
      if (!Array.isArray(parsed) || parsed.length === 0) {
        return false;
      }
      const first = parsed[0];
      if (!first || typeof first !== 'object') {
        return false;
      }
      return typeof (first as IPnpFilterGroup).filterName === 'string';
    } catch {
      return false;
    }
  }

  /**
   * Some tenants encode refiner JSON under a key other than `f`, or only in the hash.
   * Scan all query values for an array payload that looks like PnP refiners.
   */
  /**
   * PnP often uses `f` or `f_<webPartGuid>` (e.g. `f_574443da-2e4b-4967-bbb5-f425b79bc2c8`) for the same JSON payload.
   * Prefer `f`, then any `f_<guid>` key so export matches the Filters web part URL.
   */
  private _findPnpFQueryParamFromUrl(): { raw: string; key: string } | undefined {
    const entries = this._collectAllUrlParamEntries();
    const fGuidKey = /^f_[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const candidates: Array<{ key: string; value: string; rank: number }> = [];
    for (let i = 0; i < entries.length; i++) {
      const key = entries[i].key.trim();
      const value = (entries[i].value || '').trim();
      if (!value) {
        continue;
      }
      if (key === 'f') {
        candidates.push({ key, value, rank: 0 });
      } else if (fGuidKey.test(key)) {
        candidates.push({ key, value, rank: 1 });
      }
    }
    candidates.sort((a, b) => a.rank - b.rank);
    for (let c = 0; c < candidates.length; c++) {
      const raw = candidates[c].value;
      const decoded = this._decodePnpRefinerChain(raw);
      if (this._isPnpRefinerFilterJson(decoded)) {
        return { raw, key: candidates[c].key };
      }
    }
    return undefined;
  }

  private _discoverPnpRefinersInUrlParams(): { raw: string; fromKey: string } | undefined {
    const entries = this._collectAllUrlParamEntries();
    for (let i = 0; i < entries.length; i++) {
      const decoded = this._decodePnpRefinerChain(entries[i].value);
      if (this._isPnpRefinerFilterJson(decoded)) {
        return { raw: entries[i].value, fromKey: entries[i].key };
      }
    }
    return undefined;
  }

  /**
   * When PnP does not sync refiners to the URL, infer FileType FQL from a visible File type control (heuristic).
   */
  private _tryRefinementFromVisibleFilterUi(): { refinementFql: string; summary: string } | undefined {
    const hostSelector =
      'div[role="combobox"], button[aria-haspopup="listbox"], div[aria-expanded]';
    const roots = this._getPageContentRoots();
    for (let r = 0; r < roots.length; r++) {
      const hosts = roots[r].querySelectorAll(hostSelector);
      for (let i = 0; i < hosts.length; i++) {
      const host = hosts[i] as HTMLElement;
      const mount =
        host.closest('[class*="refinement"]') ||
        host.closest('[class*="Refinement"]') ||
        host.closest('[data-sp-fre-refiner]') ||
        host.closest('section');
      if (!mount) {
        continue;
      }
      const ctx = (mount.textContent || '').slice(0, 400);
      if (!/file\s*type|filetype/i.test(ctx)) {
        continue;
      }

      const pick =
        host.querySelector('.ms-Dropdown-title, span[class*="title"], [class*="dropdownLabel"]') || host;
      let val = (pick.textContent || '').trim();
      if (!val || /^(select|choose|\.\.\.)$/i.test(val)) {
        continue;
      }
      val = val.split(/\n/)[0].trim();
      const ext = val.replace(/^\./, '').split(/[\s,]+/)[0];
      if (ext && /^[a-z0-9]+$/i.test(ext) && ext.length >= 2 && ext.length <= 15) {
        const low = ext.toLowerCase();
        const r = `FileType:equals("${this._escapeFqlEqualsArg(low)}")`;
        return {
          refinementFql: r,
          summary: `${strings.FiltersFromUiLabel}: ${r}`
        };
      }
    }
    }

    for (let r = 0; r < roots.length; r++) {
    const labels = roots[r].querySelectorAll('label, span');
    for (let j = 0; j < labels.length; j++) {
      const el = labels[j] as HTMLElement;
      const t = (el.textContent || '').trim();
      if (!/^file\s*type$/i.test(t) && !/^filetype$/i.test(t)) {
        continue;
      }
      const container = el.closest('div')?.parentElement;
      if (!container) {
        continue;
      }
      const combo = container.querySelector('[role="combobox"], [role="listbox"], button');
      if (!combo) {
        continue;
      }
      const title = combo.querySelector('.ms-Dropdown-title, span') || combo;
      let val = (title.textContent || '').trim().split(/\n/)[0];
      val = val.replace(/^\./, '').trim();
      if (val && /^[a-z0-9]+$/i.test(val) && val.length >= 2 && val.length <= 15) {
        const r = `FileType:equals("${this._escapeFqlEqualsArg(val.toLowerCase())}")`;
        return {
          refinementFql: r,
          summary: `${strings.FiltersFromUiLabel}: ${r}`
        };
      }
    }
    }

    return undefined;
  }

  private _escapeFqlEqualsArg(value: string): string {
    return (value || '')
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"');
  }

  /**
   * RefinementFilters use FAST FQL. SQL-style `(a OR b)` is invalid; use `or(a, b)` / `and(a, b)`.
   */
  private _combineFqlRefinementParts(parts: string[], join: 'or' | 'and'): string {
    if (parts.length === 0) {
      return '';
    }
    if (parts.length === 1) {
      return parts[0];
    }
    return `${join}(${parts.join(', ')})`;
  }

  private _extractPnpFilterDisplayValue(_filterName: string, fv: IPnpFilterValue): string {
    let display = (fv.name || '').trim();
    if (!display && fv.value) {
      display = this._tokenFromPnpFilterValueField(fv.value);
    }
    return display;
  }

  /** Normalize ISO date string for FQL datetime("…") operands. */
  private _normalizeIsoForFqlDateTime(raw: string): string {
    let t = (raw || '').trim().replace(/^["']|["']$/g, '');
    if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
      t = `${t}T00:00:00.000Z`;
    }
    return t;
  }

  private _fqlDatetimeOperandFromIso(raw: string): string {
    const t = this._normalizeIsoForFqlDateTime(raw);
    return `datetime("${this._escapeFqlEqualsArg(t)}")`;
  }

  /** PnP `f` JSON may omit operators or use string names; normalize for Geq/Leq pairing. */
  private _coercePnpComparisonOperator(raw: unknown): number | undefined {
    if (raw === undefined || raw === null) {
      return undefined;
    }
    if (typeof raw === 'number') {
      if (isNaN(raw)) {
        return undefined;
      }
      return raw;
    }
    if (typeof raw === 'string') {
      const t = raw.trim();
      if (!t) {
        return undefined;
      }
      const n = Number(t);
      if (!isNaN(n) && String(n) === t) {
        return n;
      }
      const byName: Record<string, number> = {
        eq: PnpFilterComparisonOperator.Eq,
        neq: PnpFilterComparisonOperator.Neq,
        gt: PnpFilterComparisonOperator.Gt,
        lt: PnpFilterComparisonOperator.Lt,
        geq: PnpFilterComparisonOperator.Geq,
        leq: PnpFilterComparisonOperator.Leq,
        contains: PnpFilterComparisonOperator.Contains
      };
      const hit = byName[t.toLowerCase()];
      if (hit !== undefined) {
        return hit;
      }
    }
    return undefined;
  }

  /**
   * When date bounds were joined with `or(...)` we emit half-open FQL:
   * `range(min, hi, LE)` and `range(lo, max, GE)`. That OR is almost always true — merge to one interval.
   */
  private _mergeOrDateHalfRangeFqlIfApplicable(
    managedProperty: string,
    parts: string[],
    innerJoin: string
  ): string | undefined {
    if (String(innerJoin).toLowerCase() !== 'or' || parts.length !== 2) {
      return undefined;
    }

    const lePrefix = `${managedProperty}:range(min, datetime("`;
    const leSuffix = '"), to="LE")';
    const gePrefix = `${managedProperty}:range(datetime("`;
    const geSuffix = '"), max, from="GE")';

    const parseLeHalfIso = (s: string): string | undefined => {
      if (s.indexOf(lePrefix) !== 0 || s.lastIndexOf(leSuffix) !== s.length - leSuffix.length) {
        return undefined;
      }
      const iso = s.slice(lePrefix.length, s.length - leSuffix.length);
      return iso.indexOf('"') === -1 && iso.length > 0 ? iso : undefined;
    };

    const parseGeHalfIso = (s: string): string | undefined => {
      if (s.indexOf(gePrefix) !== 0 || s.lastIndexOf(geSuffix) !== s.length - geSuffix.length) {
        return undefined;
      }
      const iso = s.slice(gePrefix.length, s.length - geSuffix.length);
      return iso.indexOf('"') === -1 && iso.length > 0 ? iso : undefined;
    };

    const pick = (a: string, b: string): string | undefined => {
      const isoHi = parseLeHalfIso(a);
      const isoLo = parseGeHalfIso(b);
      if (!isoHi || !isoLo) {
        return undefined;
      }
      let loIso = isoLo;
      let hiIso = isoHi;
      if (this._normalizeIsoForFqlDateTime(loIso) > this._normalizeIsoForFqlDateTime(hiIso)) {
        const swap = loIso;
        loIso = hiIso;
        hiIso = swap;
      }
      return `${managedProperty}:range(${this._fqlDatetimeOperandFromIso(loIso)}, ${this._fqlDatetimeOperandFromIso(
        hiIso
      )}, from="GE", to="LE")`;
    };
    return pick(parts[0], parts[1]) ?? pick(parts[1], parts[0]);
  }

  /** ISO timestamps from PnP `f` JSON for date refiners (values[].value + operator). */
  private _collectPnpDateFilterValues(group: IPnpFilterGroup): Array<{ op: number; iso: string }> {
    const vals = group.values || [];
    const dated: Array<{ op: number; iso: string }> = [];
    for (let i = 0; i < vals.length; i++) {
      const raw = (vals[i].value || '').trim();
      if (!raw || !this._looksLikeIsoDateTimeForKql(raw)) {
        continue;
      }
      const coerced = this._coercePnpComparisonOperator(vals[i].operator);
      const op = coerced !== undefined ? coerced : PnpFilterComparisonOperator.Eq;
      dated.push({ op, iso: raw });
    }
    return dated;
  }

  /**
   * PnP date refiners belong in RefinementFilters (FQL), not KQL Querytext — same path as FileType in the refiners web part.
   * Uses RANGE / equals patterns from MS FQL reference.
   */
  private _buildDateRefinerFqlGroup(group: IPnpFilterGroup, managedProperty: string): string {
    const dated = this._collectPnpDateFilterValues(group);
    if (dated.length === 0) {
      return '';
    }

    type Dated = { op: number; iso: string };

    const innerJoin = String(group.operator || 'or').toLowerCase() === 'and' ? 'and' : 'or';

    /**
     * PnP often sets `group.operator` to **"or"** on date refiner buckets, but a **Geq + Leq** pair is still one
     * inclusive interval (AND). If we OR two half-ranges (`>= start` OR `<= end`), almost all dates match →
     * count blows past Search Results. Always fold Geq+Leq into a single RANGE regardless of group.operator.
     */
    if (dated.length === 2) {
      let geD: Dated | undefined;
      let leD: Dated | undefined;
      for (let d = 0; d < dated.length; d++) {
        const item = dated[d];
        if (item.op === PnpFilterComparisonOperator.Geq) {
          geD = item;
        }
        if (item.op === PnpFilterComparisonOperator.Leq) {
          leD = item;
        }
      }
      if (geD && leD) {
        let geIso = geD.iso;
        let leIso = leD.iso;
        if (this._normalizeIsoForFqlDateTime(geIso) > this._normalizeIsoForFqlDateTime(leIso)) {
          const swap = geIso;
          geIso = leIso;
          leIso = swap;
        }
        return `${managedProperty}:range(${this._fqlDatetimeOperandFromIso(geIso)}, ${this._fqlDatetimeOperandFromIso(
          leIso
        )}, from="GE", to="LE")`;
      }
    }

    /**
     * PnP "between" often sends two bounds both as Eq — `and(prop:"d1", prop:"d2")` would match nothing.
     * Fold into one inclusive range.
     */
    if (dated.length === 2) {
      const a = dated[0];
      const b = dated[1];
      if (a.op === PnpFilterComparisonOperator.Eq && b.op === PnpFilterComparisonOperator.Eq) {
        let lo = a.iso;
        let hi = b.iso;
        if (this._normalizeIsoForFqlDateTime(lo) > this._normalizeIsoForFqlDateTime(hi)) {
          const t = lo;
          lo = hi;
          hi = t;
        }
        return `${managedProperty}:range(${this._fqlDatetimeOperandFromIso(lo)}, ${this._fqlDatetimeOperandFromIso(
          hi
        )}, from="GE", to="LE")`;
      }
    }

    const parts: string[] = [];
    for (let j = 0; j < dated.length; j++) {
      const { op, iso } = dated[j];
      const dt = this._fqlDatetimeOperandFromIso(iso);
      switch (op) {
        case PnpFilterComparisonOperator.Eq:
          parts.push(`${managedProperty}:range(${dt}, ${dt}, from="GE", to="LE")`);
          break;
        case PnpFilterComparisonOperator.Gt:
          parts.push(`${managedProperty}:range(${dt}, max, from="GT")`);
          break;
        case PnpFilterComparisonOperator.Lt:
          parts.push(`${managedProperty}:range(min, ${dt}, to="LT")`);
          break;
        case PnpFilterComparisonOperator.Geq:
          parts.push(`${managedProperty}:range(${dt}, max, from="GE")`);
          break;
        case PnpFilterComparisonOperator.Leq:
          parts.push(`${managedProperty}:range(min, ${dt}, to="LE")`);
          break;
        case PnpFilterComparisonOperator.Neq:
          parts.push(`not(${managedProperty}:equals(${dt}))`);
          break;
        default:
          parts.push(`${managedProperty}:range(${dt}, ${dt}, from="GE", to="LE")`);
      }
    }
    if (parts.length === 0) {
      return '';
    }
    if (parts.length === 1) {
      return parts[0];
    }
    const mergedHalf = this._mergeOrDateHalfRangeFqlIfApplicable(managedProperty, parts, innerJoin);
    if (mergedHalf) {
      return mergedHalf;
    }
    return innerJoin === 'and' ? `and(${parts.join(', ')})` : `or(${parts.join(', ')})`;
  }

  /**
   * KQL for date managed properties (e.g. ModifiedOWSDATE) when not using RefinementFilters FQL.
   * Same range rules as `_buildDateRefinerFqlGroup` — two Eq bounds → `>= lo AND <= hi`.
   */
  private _buildDateKqlFilterGroup(group: IPnpFilterGroup): string {
    const managedProperty = (group.filterName || '').trim();
    if (!managedProperty) {
      return '';
    }
    const dated = this._collectPnpDateFilterValues(group);
    if (dated.length === 0) {
      return '';
    }

    type Dated = { op: number; iso: string };
    const innerJoin = String(group.operator || 'or').toLowerCase() === 'and' ? 'and' : 'or';

    if (dated.length === 2) {
      let geD: Dated | undefined;
      let leD: Dated | undefined;
      for (let d = 0; d < dated.length; d++) {
        const item = dated[d];
        if (item.op === PnpFilterComparisonOperator.Geq) {
          geD = item;
        }
        if (item.op === PnpFilterComparisonOperator.Leq) {
          leD = item;
        }
      }
      if (geD && leD) {
        let geIso = geD.iso;
        let leIso = leD.iso;
        if (this._normalizeIsoForFqlDateTime(geIso) > this._normalizeIsoForFqlDateTime(leIso)) {
          const swap = geIso;
          geIso = leIso;
          leIso = swap;
        }
        const litGe = this._isoValueToKqlDateTimeLiteral(geIso);
        const litLe = this._isoValueToKqlDateTimeLiteral(leIso);
        return `${managedProperty}>=${litGe} AND ${managedProperty}<=${litLe}`;
      }
    }

    if (dated.length === 2) {
      const a = dated[0];
      const b = dated[1];
      if (a.op === PnpFilterComparisonOperator.Eq && b.op === PnpFilterComparisonOperator.Eq) {
        let lo = a.iso;
        let hi = b.iso;
        if (this._normalizeIsoForFqlDateTime(lo) > this._normalizeIsoForFqlDateTime(hi)) {
          const t = lo;
          lo = hi;
          hi = t;
        }
        const litLo = this._isoValueToKqlDateTimeLiteral(lo);
        const litHi = this._isoValueToKqlDateTimeLiteral(hi);
        return `${managedProperty}>=${litLo} AND ${managedProperty}<=${litHi}`;
      }
    }

    const parts: string[] = [];
    for (let j = 0; j < dated.length; j++) {
      const { op, iso } = dated[j];
      const lit = this._isoValueToKqlDateTimeLiteral(iso);
      switch (op) {
        case PnpFilterComparisonOperator.Eq:
          parts.push(`${managedProperty}=${lit}`);
          break;
        case PnpFilterComparisonOperator.Neq:
          parts.push(`NOT (${managedProperty}=${lit})`);
          break;
        case PnpFilterComparisonOperator.Gt:
          parts.push(`${managedProperty}>${lit}`);
          break;
        case PnpFilterComparisonOperator.Lt:
          parts.push(`${managedProperty}<${lit}`);
          break;
        case PnpFilterComparisonOperator.Geq:
          parts.push(`${managedProperty}>=${lit}`);
          break;
        case PnpFilterComparisonOperator.Leq:
          parts.push(`${managedProperty}<=${lit}`);
          break;
        default:
          parts.push(`${managedProperty}=${lit}`);
      }
    }
    if (parts.length === 0) {
      return '';
    }
    if (parts.length === 1) {
      return parts[0];
    }
    const joiner = innerJoin === 'and' ? ' AND ' : ' OR ';
    return `(${parts.join(joiner)})`;
  }

  /** Build SharePoint FQL for a PnP FileType / FileExtension refiner group (matches search UI refiners). */
  private _buildFileTypeFqlGroup(group: IPnpFilterGroup): string {
    const vals = group.values || [];
    const pieces: string[] = [];
    for (let v = 0; v < vals.length; v++) {
      const display = this._extractPnpFilterDisplayValue((group.filterName || '').trim(), vals[v]);
      if (!display) {
        continue;
      }
      const ext = display.replace(/^\./, '').trim();
      if (!ext) {
        continue;
      }
      pieces.push(`FileType:equals("${this._escapeFqlEqualsArg(ext.toLowerCase())}")`);
    }
    if (pieces.length === 0) {
      return '';
    }
    if (pieces.length === 1) {
      return pieces[0];
    }
    const innerOpLower = String(group.operator || 'or').toLowerCase();
    return this._combineFqlRefinementParts(pieces, innerOpLower === 'and' ? 'and' : 'or');
  }

  /**
   * People / author checkbox refiners (User, Author, DisplayAuthor, …) must use RefinementFilters FQL on
   * `Author` — same as the Filters web part. KQL on `DisplayAuthor` does not match that path and skews counts.
   */
  private _buildAuthorRefinerFqlGroup(group: IPnpFilterGroup): string {
    const vals = group.values || [];
    const pieces: string[] = [];
    const filterLabel = (group.filterName || '').trim();
    for (let v = 0; v < vals.length; v++) {
      const display = this._extractPnpFilterDisplayValue(filterLabel, vals[v]);
      if (!display) {
        continue;
      }
      const esc = this._escapeFqlEqualsArg(display.trim());
      pieces.push(`Author:equals("${esc}")`);
    }
    if (pieces.length === 0) {
      return '';
    }
    if (pieces.length === 1) {
      return pieces[0];
    }
    const innerOpLower = String(group.operator || 'or').toLowerCase();
    return this._combineFqlRefinementParts(pieces, innerOpLower === 'and' ? 'and' : 'or');
  }

  /** filterName values PnP uses for “people” refiners (checkbox / user picker). */
  private _isPersonAuthorRefinerFilterName(fn: string): boolean {
    const s = (fn || '').trim().toLowerCase();
    if (!s) {
      return false;
    }
    const known = new Set<string>([
      'user',
      'users',
      'author',
      'authors',
      'displayauthor',
      'authorowsuser',
      'authorname',
      'documentauthor',
      'publisher',
      'createdby',
      'modifiedby',
      'lastmodifiedby',
      'owstaxidauthor',
      'siteauthor',
      'people',
      'person',
      'persons'
    ]);
    return known.has(s);
  }

  /**
   * PnP-managed refiner columns use managed property names like `RefinableString09`. The Filters web part
   * applies them via RefinementFilters (FQL), not QueryText — same as FileType. No per-tenant name list needed.
   */
  private _isRefinableManagedPropertyName(name: string): boolean {
    return /^Refinable(String|Date|Int|Decimal|Double|Guid)/i.test((name || '').trim());
  }

  /** `RefinableDate01` must use FQL `range()`, not `RefinableDate01:"iso"` × N (unsatisfiable AND). */
  private _isRefinableDatePropertyName(name: string): boolean {
    return /^RefinableDate/i.test((name || '').trim());
  }

  /**
   * KQL path for *OWSDATE* / *datetime* managed properties — needs `>=` / `<=`, not two `prop:"…"` ANDed.
   */
  private _isKqlDateManagedPropertyName(prop: string): boolean {
    const lower = (prop || '').trim().toLowerCase();
    if (!lower) {
      return false;
    }
    if (lower.length >= 7 && lower.substring(lower.length - 7) === 'owsdate') {
      return true;
    }
    if (lower.indexOf('datetime') >= 0) {
      return true;
    }
    return false;
  }

  /**
   * Managed property name for FQL date `range()` — maps PnP filterName (created/modified/…) to search schema names.
   */
  private _resolveDateRefinerManagedPropertyName(prop: string): string {
    const fn = (prop || '').trim().toLowerCase();
    if (fn === 'lastmodifiedtime' || fn === 'lastmodified' || fn === 'modified') {
      return 'LastModifiedTime';
    }
    if (fn === 'created') {
      return 'Created';
    }
    return (prop || '').trim();
  }

  /**
   * Date refiners: property pane `csvDateColumns` (same names as CSV date formatting) plus built-in patterns
   * (RefinableDate*, created/last modified, *owsdate*, *datetime*).
   */
  private _shouldTreatFilterAsDateRefiner(prop: string): boolean {
    const key = prop.trim().toLowerCase();
    const hints = this._getCsvExplicitDateColumns();
    if (hints && hints.has(key)) {
      return true;
    }
    if (this._isRefinableDatePropertyName(prop)) {
      return true;
    }
    if (
      key === 'created' ||
      key === 'lastmodifiedtime' ||
      key === 'lastmodified' ||
      key === 'modified'
    ) {
      return true;
    }
    if (this._isKqlDateManagedPropertyName(prop)) {
      return true;
    }
    return false;
  }

  /**
   * SharePoint often encodes refiner selections as opaque tokens in `values[].value` (e.g. leading `ǂ` + hex).
   * When present, use them in FQL equals — same data PnP sends back to search.
   */
  /** PnP sometimes wraps the token in extra JSON string quotes (`"ǂǂ…"` or leading `"` only). */
  private _unwrapPnpRefinementTokenString(raw: string): string {
    let t = (raw || '').trim();
    for (let i = 0; i < 3; i++) {
      if (t.charAt(0) === '"') {
        t = t.slice(1).trim();
      }
      if (t.charAt(t.length - 1) === '"') {
        t = t.slice(0, -1).trim();
      }
    }
    t = t.replace(/\\"/g, '"');
    return t.trim();
  }

  private _isLikelyRefinementTokenValue(raw: string): boolean {
    const t = this._unwrapPnpRefinementTokenString(raw);
    if (!t) {
      return false;
    }
    if (t.indexOf('ǂ') === 0) {
      return true;
    }
    if (/^[0-9a-fA-F]{8,}$/i.test(t)) {
      return true;
    }
    return false;
  }

  /**
   * Generic FQL for any managed property that matches `Refinable*` — covers most custom PnP filter fields.
   * SharePoint RefinementFilters expect `RefinerName:"RefinementToken"` (KeywordQuery / REST), not
   * `RefinerName:equals("…")` — the latter can return TotalRows=0 while the PnP UI still shows hits.
   */
  private _buildManagedPropertyRefinerFqlGroup(group: IPnpFilterGroup): string {
    const mp = (group.filterName || '').trim();
    if (!mp) {
      return '';
    }
    const vals = group.values || [];
    const pieces: string[] = [];
    for (let v = 0; v < vals.length; v++) {
      const fv = vals[v];
      const rawVal = (fv.value || '').trim();
      let operand: string | undefined;
      if (rawVal && this._isLikelyRefinementTokenValue(rawVal)) {
        operand = this._unwrapPnpRefinementTokenString(rawVal);
      } else {
        operand = this._extractPnpFilterDisplayValue(mp, fv).trim();
      }
      if (!operand) {
        continue;
      }
      const esc = this._escapeFqlEqualsArg(operand);
      pieces.push(`${mp}:"${esc}"`);
    }
    if (pieces.length === 0) {
      return '';
    }
    if (pieces.length === 1) {
      return pieces[0];
    }
    const innerOpLower = String(group.operator || 'or').toLowerCase();
    return this._combineFqlRefinementParts(pieces, innerOpLower === 'and' ? 'and' : 'or');
  }

  /** Map PnP URL `f` JSON to FQL + KQL: templates (FileType, dates, Author, Refinable*) then fallback KQL. */
  private _buildFilterPartsFromPnpFiltersJson(rawJson: string): { filterKql: string; refinementFql: string } {
    let parsed: IPnpFilterGroup[];
    try {
      parsed = JSON.parse(rawJson) as IPnpFilterGroup[];
    } catch {
      return { filterKql: '', refinementFql: '' };
    }
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return { filterKql: '', refinementFql: '' };
    }

    const kqlGroupClauses: string[] = [];
    const fqlGroupClauses: string[] = [];

    for (let g = 0; g < parsed.length; g++) {
      const group = parsed[g];
      const prop = (group.filterName || '').trim();
      if (!prop) {
        continue;
      }

      const fn = prop.toLowerCase();
      if (fn === 'filetype' || fn === 'fileextension') {
        const fg = this._buildFileTypeFqlGroup(group);
        if (fg) {
          fqlGroupClauses.push(fg);
        }
        continue;
      }

      if (this._isPersonAuthorRefinerFilterName(fn)) {
        const ag = this._buildAuthorRefinerFqlGroup(group);
        if (ag) {
          fqlGroupClauses.push(ag);
        }
        continue;
      }

      if (this._shouldTreatFilterAsDateRefiner(prop)) {
        const managedForDate = this._resolveDateRefinerManagedPropertyName(prop);
        const dg = this._buildDateRefinerFqlGroup(group, managedForDate);
        if (dg) {
          fqlGroupClauses.push(dg);
          continue;
        }
        if (this._isRefinableManagedPropertyName(prop)) {
          const rg = this._buildManagedPropertyRefinerFqlGroup(group);
          if (rg) {
            fqlGroupClauses.push(rg);
          }
          continue;
        }
        const kg = this._buildDateKqlFilterGroup(group);
        if (kg) {
          kqlGroupClauses.push(kg);
          continue;
        }
      }

      if (this._isRefinableManagedPropertyName(prop)) {
        const rg = this._buildManagedPropertyRefinerFqlGroup(group);
        if (rg) {
          fqlGroupClauses.push(rg);
        }
        continue;
      }

      const vals = group.values || [];
      const valueTokens: string[] = [];
      for (let v = 0; v < vals.length; v++) {
        const token = this._filterValueToKqlToken(prop, vals[v]);
        if (token) {
          valueTokens.push(token);
        }
      }
      if (valueTokens.length === 0) {
        continue;
      }

      const innerOpLower = String(group.operator || 'or').toLowerCase();
      const innerOp = innerOpLower === 'and' ? ' AND ' : ' OR ';
      kqlGroupClauses.push(
        valueTokens.length === 1 ? valueTokens[0] : `(${valueTokens.join(innerOp)})`
      );
    }

    const filterKql =
      kqlGroupClauses.length === 0
        ? ''
        : kqlGroupClauses.length === 1
          ? kqlGroupClauses[0]
          : `(${kqlGroupClauses.join(') AND (')})`;

    const refinementFql =
      fqlGroupClauses.length === 0
        ? ''
        : fqlGroupClauses.length === 1
          ? fqlGroupClauses[0]
          : `and(${fqlGroupClauses.join(', ')})`;

    return { filterKql, refinementFql };
  }

  /**
   * PnP Modern Search stores refiners in query string `f` as URL-encoded JSON:
   * [{"filterName":"FileType","values":[{"name":"docx",...}],"operator":"or"}]
   * Also discovers the same JSON under other param keys / hash, then falls back to visible File type UI.
   */
  private _getUrlFilterParts(): { filterKql: string; refinementFql: string; summary: string } {
    // Always apply PnP URL `f` (and discovery / UI fallback) via RefinementFilters — matches Filters web part.
    const paramKey = 'f';
    let raw: string | undefined;
    let summaryPrefix: string | undefined;

    const named = this._getUrlParam(paramKey);
    if (named && named.trim()) {
      raw = named.trim();
      summaryPrefix = `${strings.FromUrlLabel} (${paramKey})`;
    } else {
      const fScoped = this._findPnpFQueryParamFromUrl();
      if (fScoped) {
        raw = fScoped.raw.trim();
        summaryPrefix = `${strings.FromUrlLabel} (${fScoped.key})`;
      } else {
        const discovered = this._discoverPnpRefinersInUrlParams();
        if (discovered) {
          raw = discovered.raw.trim();
          summaryPrefix = `${strings.FiltersDiscoveredInUrlLabel} (${discovered.fromKey})`;
        }
      }
    }

    if (raw) {
      const decoded = this._decodePnpRefinerChain(raw);
      const parts = this._buildFilterPartsFromPnpFiltersJson(decoded);
      if ((parts.filterKql || parts.refinementFql) && summaryPrefix) {
        const bits: string[] = [];
        if (parts.refinementFql) {
          bits.push(parts.refinementFql);
        }
        if (parts.filterKql) {
          bits.push(`KQL ${parts.filterKql}`);
        }
        return {
          filterKql: parts.filterKql,
          refinementFql: parts.refinementFql,
          summary: `${summaryPrefix}: ${bits.join(' | ')}`
        };
      }
      // Invalid `f` or unexpected JSON – try DOM before giving up.
    }

    const domHit = this._tryRefinementFromVisibleFilterUi();
    if (domHit) {
      return {
        filterKql: '',
        refinementFql: domHit.refinementFql,
        summary: domHit.summary
      };
    }

    if (raw) {
      return { filterKql: '', refinementFql: '', summary: strings.FiltersParseFailedLabel };
    }

    return { filterKql: '', refinementFql: '', summary: strings.NoFiltersInUrlLabel };
  }

  /**
   * PnP date filters put a real ISO timestamp in `value` and a label in `name` (e.g. "Older than a year").
   * KQL must use `value` + `operator`, not the display name.
   */
  private _looksLikeIsoDateTimeForKql(s: string): boolean {
    const t = (s || '').trim().replace(/^["']|["']$/g, '');
    return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(t) || /^\d{4}-\d{2}-\d{2}$/.test(t);
  }

  private _isoValueToKqlDateTimeLiteral(raw: string): string {
    let t = (raw || '').trim().replace(/^["']|["']$/g, '');
    if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
      t = `${t}T00:00:00.000Z`;
    }
    // KQL datetime literal per MS docs: Created=datetime("2021-07-30T08:35:57.000Z")
    return `datetime("${t}")`;
  }

  private _tryPnpDateFilterToKql(managedProperty: string, fv: IPnpFilterValue): string {
    const rawVal = (fv.value || '').trim();
    if (!rawVal || !this._looksLikeIsoDateTimeForKql(rawVal)) {
      return '';
    }

    const lit = this._isoValueToKqlDateTimeLiteral(rawVal);
    const op =
      fv.operator !== undefined && fv.operator !== null ? Number(fv.operator) : PnpFilterComparisonOperator.Eq;

    switch (op) {
      case PnpFilterComparisonOperator.Eq:
        return `${managedProperty}=${lit}`;
      case PnpFilterComparisonOperator.Neq:
        return `NOT (${managedProperty}=${lit})`;
      case PnpFilterComparisonOperator.Gt:
        return `${managedProperty}>${lit}`;
      case PnpFilterComparisonOperator.Lt:
        return `${managedProperty}<${lit}`;
      case PnpFilterComparisonOperator.Geq:
        return `${managedProperty}>=${lit}`;
      case PnpFilterComparisonOperator.Leq:
        return `${managedProperty}<=${lit}`;
      default:
        return `${managedProperty}=${lit}`;
    }
  }

  private _filterValueToKqlToken(filterName: string, fv: IPnpFilterValue): string {
    const fn = filterName.toLowerCase();
    if (fn === 'filetype' || fn === 'fileextension') {
      const display = this._extractPnpFilterDisplayValue(filterName, fv);
      if (!display) {
        return '';
      }
      const ext = display.replace(/^\./, '');
      return `fileextension:${ext}`;
    }

    // Map common PnP refiner "filterName" values to searchable managed property names.
    let prop = filterName;
    if (fn === 'lastmodifiedtime' || fn === 'lastmodified' || fn === 'modified') {
      prop = 'LastModifiedTime';
    } else if (fn === 'created') {
      prop = 'Created';
    }

    const isDateRefiner =
      fn === 'created' || fn === 'lastmodifiedtime' || fn === 'lastmodified' || fn === 'modified';
    if (isDateRefiner) {
      const dateKql = this._tryPnpDateFilterToKql(prop, fv);
      if (dateKql) {
        return dateKql;
      }
    }

    const display = this._extractPnpFilterDisplayValue(filterName, fv);
    if (!display) {
      return '';
    }

    const escaped = display.replace(/"/g, '\\"');
    if (/[\s:]/.test(display)) {
      return `${prop}:"${escaped}"`;
    }
    return `${prop}:${escaped}`;
  }

  /** Decode hex embedded in PnP filter value strings (fallback when name is missing). */
  private _tokenFromPnpFilterValueField(raw: string): string {
    const s = String(raw).replace(/\\/g, '');
    const hexMatch = s.match(/([0-9a-fA-F]{4,})/);
    if (hexMatch && hexMatch[1] && hexMatch[1].length % 2 === 0) {
      try {
        const hex = hexMatch[1];
        let out = '';
        for (let i = 0; i < hex.length; i += 2) {
          out += String.fromCharCode(parseInt(hex.substr(i, 2), 16));
        }
        if (/^[\x20-\x7E]+$/.test(out)) {
          return out.trim();
        }
      } catch {
        // ignore
      }
    }
    return s.replace(/^["']|["']$/g, '').trim();
  }

  private _escapeCsvCell(value: string): string {
    const v = value || '';
    const escaped = v.replace(/"/g, '""');
    return /[",\r\n]/.test(v) ? `"${escaped}"` : escaped;
  }

  private _downloadCsv(fileName: string, content: string): void {
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    this.domElement.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  private _normalizeToString(value: unknown): string {
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

  /** Strip braces/quotes so Search API gets a plain GUID. */
  private _normalizeSourceId(raw: string): string {
    let s = (raw || '').trim();
    s = s.replace(/^['"]|['"]$/g, '');
    if (s.charAt(0) === '{' && s.charAt(s.length - 1) === '}') {
      s = s.slice(1, -1).trim();
    }
    return s;
  }

  /**
   * postquery SourceId must be Edm.Guid. Extract first GUID from pasted text,
   * accept 32-char hex without hyphens; return lowercase `xxxxxxxx-xxxx-...` for JSON.
   */
  private _formatSourceIdForSearchApi(raw: string): string {
    let s = this._normalizeSourceId(raw);
    if (!s) {
      throw new Error(strings.SourceIdRequiredError);
    }

    const guidWithDashes = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
    const embedded = s.match(
      /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/i
    );
    if (embedded) {
      s = embedded[0];
    } else {
      const hexOnly = s.replace(/[^0-9a-fA-F]/gi, '');
      if (hexOnly.length === 32) {
        s = `${hexOnly.slice(0, 8)}-${hexOnly.slice(8, 12)}-${hexOnly.slice(12, 16)}-${hexOnly.slice(16, 20)}-${hexOnly.slice(20, 32)}`;
      }
    }

    if (!guidWithDashes.test(s)) {
      throw new Error(strings.InvalidSourceIdGuidError);
    }

    // Edm.Guid in JSON body: unbraced lowercase (braced string often breaks conversion).
    return s.toLowerCase();
  }

  /** KQL: escape single quotes by doubling. */
  private _escapeKqlQuotes(kql: string): string {
    return (kql || '').replace(/'/g, "''");
  }

  /** First-level keys on JSON (for debug only). */
  private _jsonTopKeysForDebug(root: unknown, max: number = 12): string {
    if (!root || typeof root !== 'object' || Array.isArray(root)) return '(non-object)';
    const keys = Object.keys(root as Record<string, unknown>);
    return keys.slice(0, max).join(',') + (keys.length > max ? '…' : '');
  }

  private _getPath(root: unknown, segments: string[]): unknown {
    let cur: unknown = root;
    for (let i = 0; i < segments.length; i++) {
      if (cur === null || cur === undefined) return undefined;
      if (typeof cur !== 'object') return undefined;
      cur = (cur as Record<string, unknown>)[segments[i]];
    }
    return cur;
  }

  /**
   * SharePoint `postquery` JSON shape differs by OData mode (`verbose` vs `nometadata`) and version.
   * Walk the tree to find the RelevantResults-like block: has Table.Rows and TotalRows(-like) fields.
   */
  private _findRelevantResultsBlock(root: unknown): { block: Record<string, unknown> | undefined; how: string } {
    const tryFromPrimary = (primary: unknown, path: string): { block: Record<string, unknown>; how: string } | undefined => {
      if (!primary || typeof primary !== 'object') return undefined;
      const p = primary as Record<string, unknown>;
      const rr = p.RelevantResults;
      if (!rr || typeof rr !== 'object') return undefined;
      const rro = rr as Record<string, unknown>;
      const table = rro.Table;
      if (!table || typeof table !== 'object') return undefined;
      if ((table as Record<string, unknown>).Rows === undefined) return undefined;
      return { block: rro, how: path };
    };

    const orderedPaths: Array<{ segments: string[]; path: string }> = [
      { segments: ['d', 'postquery', 'PrimaryQueryResult'], path: 'd.postquery.PrimaryQueryResult' },
      { segments: ['d', 'PostQuery', 'PrimaryQueryResult'], path: 'd.PostQuery.PrimaryQueryResult' },
      { segments: ['d', 'query', 'PrimaryQueryResult'], path: 'd.query.PrimaryQueryResult' },
      { segments: ['postquery', 'PrimaryQueryResult'], path: 'postquery.PrimaryQueryResult' },
      { segments: ['PostQuery', 'PrimaryQueryResult'], path: 'PostQuery.PrimaryQueryResult' },
      { segments: ['value', 'postquery', 'PrimaryQueryResult'], path: 'value.postquery.PrimaryQueryResult' },
      { segments: ['value', 'PostQuery', 'PrimaryQueryResult'], path: 'value.PostQuery.PrimaryQueryResult' }
    ];

    for (let i = 0; i < orderedPaths.length; i++) {
      const primary = this._getPath(root, orderedPaths[i].segments);
      const hit = tryFromPrimary(primary, orderedPaths[i].path);
      if (hit) return hit;
    }

    const stack: unknown[] = [root];
    let visits = 0;
    const maxVisits = 10000;
    while (stack.length > 0 && visits < maxVisits) {
      const cur = stack.pop();
      visits++;
      if (cur === null || cur === undefined) continue;
      if (typeof cur !== 'object') continue;
      if (Array.isArray(cur)) {
        for (let a = 0; a < cur.length; a++) stack.push(cur[a]);
        continue;
      }
      const o = cur as Record<string, unknown>;
      const table = o.Table;
      if (table && typeof table === 'object') {
        const t = table as Record<string, unknown>;
        const hasRows = t.Rows !== undefined;
        const looksRelevant =
          hasRows && ('TotalRows' in o || 'TotalRowsIncludingDuplicates' in o || 'RowCount' in o);
        if (looksRelevant) {
          return { block: o, how: `deep-scan@visit${visits}` };
        }
      }
      const keys = Object.keys(o);
      for (let k = 0; k < keys.length; k++) stack.push(o[keys[k]]);
    }

    // Last resort: any object with Table.Rows (some payloads omit TotalRows on the same node).
    const stackWeak: unknown[] = [root];
    let visitsW = 0;
    while (stackWeak.length > 0 && visitsW < maxVisits) {
      const cur = stackWeak.pop();
      visitsW++;
      if (cur === null || cur === undefined) continue;
      if (typeof cur !== 'object') continue;
      if (Array.isArray(cur)) {
        for (let a = 0; a < cur.length; a++) stackWeak.push(cur[a]);
        continue;
      }
      const o = cur as Record<string, unknown>;
      const table = o.Table;
      if (table && typeof table === 'object') {
        const t = table as Record<string, unknown>;
        const rowsNode = t.Rows;
        let n = 0;
        if (Array.isArray(rowsNode)) n = rowsNode.length;
        else if (rowsNode && typeof rowsNode === 'object') {
          const rn = rowsNode as { results?: unknown };
          if (Array.isArray(rn.results)) n = rn.results.length;
        }
        if (n > 0) {
          return { block: o, how: `deep-scan-weak@visit${visitsW}` };
        }
      }
      const keysW = Object.keys(o);
      for (let k = 0; k < keysW.length; k++) stackWeak.push(o[keysW[k]]);
    }

    return { block: undefined, how: 'not-found' };
  }

  private async _postSearchPostquery(
    postUrl: string,
    payload: Record<string, unknown>,
    odataMode: 'nometadata' | 'verbose'
  ): Promise<unknown> {
    const accept =
      odataMode === 'verbose'
        ? 'application/json;odata=verbose;charset=utf-8'
        : 'application/json;odata=nometadata;charset=utf-8';
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(postUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: accept,
        'Content-Type': accept,
        'odata-version': ''
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      let detail = '';
      try {
        const errBody = await response.json();
        detail =
          (errBody &&
            (errBody.error?.message?.value || errBody['odata.error']?.message?.value || errBody.error?.message)) ||
          JSON.stringify(errBody).slice(0, 400);
      } catch {
        try {
          detail = (await response.text()).slice(0, 400);
        } catch {
          detail = '';
        }
      }
      throw new Error(`HTTP ${response.status}${detail ? `: ${detail}` : ''}`);
    }

    return await response.json();
  }

  /**
   * Classic Search REST: GET `/_api/search/query` returns `d.query.PrimaryQueryResult` (verbose).
   * Fallback when POST `postquery` parses as empty on some farms/builds.
   */
  private async _getSearchQueryViaGet(
    webUrl: string,
    querytext: string,
    sourceId: string,
    rowLimit: number,
    selectProps: string[],
    refinementFiltersFql?: string
  ): Promise<unknown> {
    const root = webUrl.replace(/\/$/, '');
    const inner = this._escapeKqlQuotes(querytext);
    const querytextParam = `'${inner}'`;
    const sourceParam = `'${sourceId}'`;
    const selectParam = `'${selectProps.join(',')}'`;
    const sortParam = `'DocId:ascending'`;
    let url =
      `${root}/_api/search/query` +
      `?querytext=${encodeURIComponent(querytextParam)}` +
      `&sourceid=${encodeURIComponent(sourceParam)}` +
      `&rowlimit=${encodeURIComponent(String(rowLimit))}` +
      `&selectproperties=${encodeURIComponent(selectParam)}` +
      `&sortlist=${encodeURIComponent(sortParam)}` +
      `&trimduplicates=${encodeURIComponent('false')}`;

    const rf = (refinementFiltersFql || '').trim();
    if (rf) {
      const rfInner = rf.replace(/'/g, "''");
      const refinementParam = `'${rfInner}'`;
      url += `&refinementfilters=${encodeURIComponent(refinementParam)}`;
    }

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'odata-version': ''
      }
    });

    if (!response.ok) {
      let detail = '';
      try {
        const errBody = await response.json();
        detail =
          (errBody &&
            (errBody.error?.message?.value || errBody['odata.error']?.message?.value || errBody.error?.message)) ||
          JSON.stringify(errBody).slice(0, 400);
      } catch {
        try {
          detail = (await response.text()).slice(0, 400);
        } catch {
          detail = '';
        }
      }
      throw new Error(`GET query HTTP ${response.status}${detail ? `: ${detail}` : ''}`);
    }

    return await response.json();
  }

  private _extractRowsFromSearchJson(
    json: unknown,
    exportColumnKeys: string[]
  ): {
    rows: Array<Record<string, string>>;
    lastDocId?: number;
    totalRows?: number;
    totalRowsRawType: string;
    totalRowsRawValue: string;
    tableRowsIsArray: boolean;
    tableRowsHasResultsArray: boolean;
    tableRowsResultsLength?: number;
    primaryPath: string;
    relevantDefined: boolean;
    relevantHow: string;
    jsonTopKeys: string;
  } {
    const found = this._findRelevantResultsBlock(json);
    const relevant = found.block;
    const primaryPath = found.how;
    const relevantDefined = !!relevant;
    const jsonTopKeys = this._jsonTopKeysForDebug(json);

    const tableRows: unknown =
      relevant && relevant.Table && typeof relevant.Table === 'object'
        ? (relevant.Table as Record<string, unknown>).Rows
        : undefined;
    const tableRowsIsArray = Array.isArray(tableRows);
    const tableRowsObj = typeof tableRows === 'object' && tableRows !== null ? (tableRows as { results?: unknown; Row?: unknown }) : undefined;
    const tableRowsHasResultsArray = !!(tableRowsObj && Array.isArray(tableRowsObj.results));
    const tableRowsResultsLength = tableRowsHasResultsArray ? (tableRowsObj!.results as Array<unknown>).length : undefined;

    const rows = (() => {
      if (Array.isArray(tableRows)) return tableRows;
      if (tableRows && typeof tableRows === 'object') {
        const obj = tableRows as { results?: unknown; Row?: unknown };
        if (Array.isArray(obj.results)) return obj.results;
        if (Array.isArray(obj.Row)) return obj.Row;
      }
      return [];
    })();

    const toCellsArray = (cellsContainer: unknown): Array<{ Key?: unknown; Value?: unknown }> => {
      if (!cellsContainer) return [];
      if (Array.isArray(cellsContainer)) {
        return cellsContainer as Array<{ Key?: unknown; Value?: unknown }>;
      }
      if (cellsContainer && typeof cellsContainer === 'object') {
        const obj = cellsContainer as { results?: unknown };
        if (Array.isArray(obj.results)) {
          return obj.results as Array<{ Key?: unknown; Value?: unknown }>;
        }
      }
      return [];
    };

    const exportColumnLower: string[] = [];
    for (let c = 0; c < exportColumnKeys.length; c++) {
      exportColumnLower.push(exportColumnKeys[c].toLowerCase());
    }

    const csvDateHints = this._getCsvExplicitDateColumns();

    const mapped = rows.map((row) => {
      const rowObj = row as { Cells?: unknown };
      const cells = toCellsArray(rowObj?.Cells);
      const prepared = prepareSearchRowCells(cells, (v: unknown) => this._normalizeToString(v));
      const out: Record<string, string> = {};
      for (let k = 0; k < exportColumnKeys.length; k++) {
        const name = exportColumnKeys[k];
        const cell = getPreparedCellValueForColumn(prepared, exportColumnLower[k]);
        out[name] = formatCsvDateCell(name, cell, csvDateHints);
      }
      return out;
    });

    const lastRow = rows.length > 0 ? rows[rows.length - 1] : undefined;
    const lastRowObj = lastRow as { Cells?: unknown } | undefined;
    const lastCells = toCellsArray(lastRowObj?.Cells);
    const lastPrepared = prepareSearchRowCells(lastCells, (v: unknown) => this._normalizeToString(v));
    const lastDocIdRaw = getPreparedCellValueForCandidates(lastPrepared, ['IndexDocId', 'indexdocid', 'DocId', 'docid']);
    const lastDocId = lastDocIdRaw ? Number(lastDocIdRaw) : undefined;

    const totalRowsRaw = relevant?.TotalRows as unknown;
    const totalRowsRawType = totalRowsRaw === null ? 'null' : typeof totalRowsRaw;
    const totalRowsRawValue = this._formatUnknownForDebug(totalRowsRaw);
    const totalRowsTyped = typeof totalRowsRaw === 'number' ? totalRowsRaw : undefined;

    return {
      rows: mapped,
      lastDocId: typeof lastDocId === 'number' && !isNaN(lastDocId) ? lastDocId : undefined,
      totalRows: totalRowsTyped,
      totalRowsRawType,
      totalRowsRawValue,
      tableRowsIsArray,
      tableRowsHasResultsArray,
      tableRowsResultsLength,
      primaryPath,
      relevantDefined,
      relevantHow: found.how,
      jsonTopKeys
    };
  }

  private async _fetchExportPage(params: {
    webUrl: string;
    sourceId: string;
    queryText: string;
    pageSize: number;
    selectProperties: string;
    selectPropertiesList?: string[];
    /** Managed property names for CSV (subset of SelectProperties). */
    exportColumnKeys: string[];
    /** FQL refinement strings (e.g. `FileType:equals("docx")`) — same mechanism as PnP filter web part. */
    refinementFilters?: string[];
    /** If false, skip GET `search/query` when postquery returns 0 rows (IndexDocId pages: empty = done). */
    enableGetFallbackWhenEmpty?: boolean;
  }): Promise<{
    rows: Array<Record<string, string>>;
    lastDocId?: number;
    totalRows?: number;
    debug: {
      sentQueryText: string;
      sentRefinementFilters: string;
      sentSourceId: string;
      extractedRows: number;
      totalRowsRawType: string;
      totalRowsRawValue: string;
      tableRowsIsArray: boolean;
      tableRowsHasResultsArray: boolean;
      tableRowsResultsLength?: number;
      primaryPath: string;
      relevantDefined: boolean;
      relevantHow: string;
      odataAttempt: string;
      jsonTopKeys: string;
      transport: string;
    };
  }> {
    const webUrl = params.webUrl.replace(/\/$/, '');
    const postUrl = `${webUrl}/_api/search/postquery`;
    const safeQuery = this._escapeKqlQuotes(params.queryText);
    const sourceId = this._formatSourceIdForSearchApi(params.sourceId);
    const selectProps =
      params.selectPropertiesList && params.selectPropertiesList.length > 0
        ? params.selectPropertiesList
        : params.selectProperties
            .split(',')
            .map((s) => s.trim())
            .filter((s) => s.length > 0);

    const refinementList =
      params.refinementFilters !== undefined
        ? params.refinementFilters.map((s) => (s || '').trim()).filter((s) => s.length > 0)
        : [];
    const refinementFqlForGet =
      refinementList.length === 0
        ? undefined
        : refinementList.length === 1
          ? refinementList[0]
          : `and(${refinementList.join(', ')})`;

    // postquery expects JSON arrays here, not OData-style { results: [...] } (causes 400 StartArray vs StartObject).
    const requestBody: Record<string, unknown> = {
      Querytext: safeQuery,
      RowLimit: params.pageSize,
      RowsPerPage: params.pageSize,
      SelectProperties: selectProps,
      SortList: [{ Property: 'DocId', Direction: 0 }],
      SourceId: sourceId,
      TrimDuplicates: false,
      /** Faster export: skip query rules / ranking extras the results UI may use. */
      EnableQueryRules: false
    };
    if (refinementList.length > 0) {
      requestBody.RefinementFilters = refinementList;
    }
    const payload = { request: requestBody };

    // Prefer nometadata first: smaller payloads and faster JSON parse than verbose (major win on large RowLimit).
    let json: unknown;
    let odataAttempt = 'nometadata';
    try {
      json = await this._postSearchPostquery(postUrl, payload, 'nometadata');
    } catch (firstErr) {
      try {
        odataAttempt = 'verbose(fallback)';
        json = await this._postSearchPostquery(postUrl, payload, 'verbose');
      } catch {
        throw firstErr instanceof Error ? firstErr : new Error(String(firstErr));
      }
    }

    const rowCountFromBlock = (b: Record<string, unknown> | undefined): number => {
      if (!b) return 0;
      const tr = b.Table as Record<string, unknown> | undefined;
      if (!tr || tr.Rows === undefined) return 0;
      const raw = tr.Rows as unknown;
      if (Array.isArray(raw)) return raw.length;
      if (raw && typeof raw === 'object') {
        const rr = raw as { results?: unknown; Row?: unknown };
        if (Array.isArray(rr.results)) return rr.results.length;
        if (Array.isArray(rr.Row)) return rr.Row.length;
      }
      return 0;
    };

    let found = this._findRelevantResultsBlock(json);
    if (!found.block || rowCountFromBlock(found.block) === 0) {
      try {
        const altMode = odataAttempt === 'nometadata' ? 'verbose' : 'nometadata';
        const json2 = await this._postSearchPostquery(postUrl, payload, altMode);
        const found2 = this._findRelevantResultsBlock(json2);
        if (found2.block && (rowCountFromBlock(found2.block) > 0 || !found.block)) {
          json = json2;
          found = found2;
          odataAttempt = `${odataAttempt}+${altMode}`;
        }
      } catch {
        // keep first json
      }
    }

    const colKeys = params.exportColumnKeys && params.exportColumnKeys.length > 0 ? params.exportColumnKeys : ['Title', 'Path', 'Author'];
    let transport = `postquery:${odataAttempt}`;
    let ex = this._extractRowsFromSearchJson(json, colKeys);

    // Some tenants accept RefinementFilters as a single FQL string; array form can yield 0 rows.
    if (ex.rows.length === 0 && refinementList.length === 1) {
      const requestBodyStr: Record<string, unknown> = { ...requestBody, RefinementFilters: refinementList[0] };
      const payloadStr = { request: requestBodyStr };
      try {
        let jsonStr: unknown;
        try {
          jsonStr = await this._postSearchPostquery(postUrl, payloadStr, 'nometadata');
        } catch {
          jsonStr = await this._postSearchPostquery(postUrl, payloadStr, 'verbose');
        }
        const exStr = this._extractRowsFromSearchJson(jsonStr, colKeys);
        if (exStr.rows.length > 0) {
          json = jsonStr;
          ex = exStr;
          transport = `${transport};refinementFiltersAsString`;
        }
      } catch {
        // keep first result
      }
    }

    // With refiners only, some tenants return TotalRows=0 when Querytext is `*`; empty matches "all" + refinement.
    if (ex.rows.length === 0 && refinementList.length > 0 && (params.queryText || '').trim() === '*') {
      const requestBodyNoText: Record<string, unknown> = { ...requestBody, Querytext: '' };
      const payloadNoText = { request: requestBodyNoText };
      try {
        let jsonNoText: unknown;
        try {
          jsonNoText = await this._postSearchPostquery(postUrl, payloadNoText, 'nometadata');
        } catch {
          jsonNoText = await this._postSearchPostquery(postUrl, payloadNoText, 'verbose');
        }
        const exNoText = this._extractRowsFromSearchJson(jsonNoText, colKeys);
        if (exNoText.rows.length > 0) {
          json = jsonNoText;
          ex = exNoText;
          transport = `${transport};querytextEmptyWithRefiners`;
        }
      } catch {
        // keep prior extraction
      }
    }

    const allowGetFallback = params.enableGetFallbackWhenEmpty !== false;
    if (ex.rows.length === 0 && allowGetFallback) {
      try {
        const jsonGet = await this._getSearchQueryViaGet(
          webUrl,
          safeQuery,
          sourceId,
          params.pageSize,
          selectProps,
          refinementFqlForGet
        );
        const exGet = this._extractRowsFromSearchJson(jsonGet, colKeys);
        if (exGet.rows.length > 0) {
          ex = exGet;
          transport = 'GET /_api/search/query';
        }
      } catch {
        // keep postquery extraction (likely still empty)
      }
    }

    return {
      rows: ex.rows,
      lastDocId: ex.lastDocId,
      totalRows: ex.totalRows,
      debug: {
        sentQueryText: params.queryText,
        sentRefinementFilters: refinementList.length > 0 ? refinementList.join(' | ') : '',
        sentSourceId: sourceId,
        extractedRows: ex.rows.length,
        totalRowsRawType: ex.totalRowsRawType,
        totalRowsRawValue: ex.totalRowsRawValue,
        tableRowsIsArray: ex.tableRowsIsArray,
        tableRowsHasResultsArray: ex.tableRowsHasResultsArray,
        tableRowsResultsLength: ex.tableRowsResultsLength,
        primaryPath: ex.primaryPath,
        relevantDefined: ex.relevantDefined,
        relevantHow: ex.relevantHow,
        odataAttempt,
        jsonTopKeys: ex.jsonTopKeys,
        transport
      }
    };
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: '',
              groupFields: [
                PropertyPaneTextField('sourceId', {
                  label: strings.SourceIdLabel
                }),
                PropertyPaneTextField('exportColumns', {
                  label: strings.ExportColumnsLabel,
                  description: strings.ExportColumnsDescription,
                  multiline: true
                }),
                PropertyPaneTextField('csvDateColumns', {
                  label: strings.CsvDateColumnsLabel,
                  description: strings.CsvDateColumnsDescription,
                  multiline: true
                }),
                PropertyPaneToggle('debugApi', {
                  label: strings.DebugApiLabel
                })
              ]
            },
            {
              groupName: strings.ButtonAppearanceGroupLabel,
              groupFields: buildExportButtonAppearanceGroupFields(strings)
            }
          ]
        }
      ]
    };
  }
}
