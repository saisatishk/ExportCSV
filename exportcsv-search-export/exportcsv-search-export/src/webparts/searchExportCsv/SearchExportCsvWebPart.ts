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

/** Dispatched when `history.pushState` / `replaceState` run (PnP Modern Search updates filters this way). */
const SEARCH_EXPORT_LOCATION_CHANGE = 'searchExportCsvLocationChange';

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

export interface ISearchExportCsvWebPartProps {
  /** Result source GUID (must match Search Results). Optional URL override: `sourceid`. */
  sourceId: string;
  /**
   * When true, read PnP Filters URL `f` JSON (and same discovery as before) and apply via
   * SharePoint RefinementFilters (FQL) so results match the Filters + Search Results web parts.
   */
  appendUrlFilters: boolean;
  /** Show API extraction diagnostics in the UI (useful while debugging locally). */
  debugApi?: boolean;
}

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

/**
 * Mirrors PnP `FilterComparisonOperator` (see `@pnp/modern-search-extensibility` IDataFilter.d.ts).
 * URL `f` JSON stores the numeric enum on each filter value.
 */
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
      this.properties.appendUrlFilters === true &&
      (!!((filterParts.refinementFql || '').trim()) || !!((filterParts.filterKql || '').trim()));

    const noKeywordsHintMessage = !effectiveQueryText.value.trim()
      ? hasActiveRefinersForExport
        ? strings.ExportNoKeywordsWithRefinersHint
        : strings.ExportNoKeywordsNoRefinersHint
      : '';

    this.domElement.innerHTML = `
      <section class="${styles.searchExportCsv}">
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
        </div>

        <div class="${styles.actions}">
          <button type="button" class="${styles.button}" data-action="export">${strings.ExportButtonLabel}</button>
          <button type="button" class="${styles.button}" data-action="cancel" disabled>${strings.CancelButtonLabel}</button>
        </div>

        <div class="${styles.status}" data-role="status"></div>
      </section>
    `;

    const exportButton = this.domElement.querySelector<HTMLButtonElement>('button[data-action="export"]');
    const cancelButton = this.domElement.querySelector<HTMLButtonElement>('button[data-action="cancel"]');
    const status = this.domElement.querySelector<HTMLDivElement>('div[data-role="status"]');

    if (!exportButton || !cancelButton || !status) return;

    const showError = (message: string): void => {
      status.textContent = message;
      status.className = `${styles.status} ${styles.error}`;
      exportButton.disabled = false;
      cancelButton.disabled = true;
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
      cancelButton.disabled = false;
      status.textContent = strings.ExportStarted;
      status.className = `${styles.status}`;

      try {
        const columns = ['Title', 'Path', 'Author'];
        const selectProperties = ['Title', 'Path', 'Author', 'DocId'].join(',');
        const pageSize = 500;
        const maxRows = 200000;
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

        while (shouldContinue) {
          if (this._isCancelled) {
            status.textContent = strings.ExportCancelled;
            return;
          }

          pageIndex++;
          const hasBaseQuery = queryBase.trim().length > 0;
          const effectiveQuery =
            lastDocId === undefined
              ? queryBase
              : hasBaseQuery
                ? (queryBase === '*' ? `IndexDocId>${lastDocId}` : `(${queryBase}) AND IndexDocId>${lastDocId}`)
                : `IndexDocId>${lastDocId}`;

          const result = await this._fetchExportPage({
            webUrl: this.context.pageContext.web.absoluteUrl,
            sourceId,
            queryText: effectiveQuery,
            pageSize,
            selectProperties,
            refinementFilters: refinementFiltersPayload
          });

          if (totalRows === undefined && result.totalRows !== undefined) {
            totalRows = result.totalRows;
          }
          lastDebug = result.debug;

          for (let i = 0; i < result.rows.length; i++) {
            const row = result.rows[i];
            csvLines.push(
              [
                this._escapeCsvCell(row.Title),
                this._escapeCsvCell(row.Path),
                this._escapeCsvCell(row.Author)
              ].join(',')
            );
          }

          exported += result.rows.length;
          status.textContent = `${strings.ExportInProgress} ${exported}${totalRows ? ` / ${totalRows}` : ''} (${strings.PageLabel} ${pageIndex})`;

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
        const message = error instanceof Error ? error.message : String(error);
        showError(`${strings.ExportFailedPrefix} ${message}`);
      } finally {
        exportButton.disabled = false;
        cancelButton.disabled = true;
      }
    };

    cancelButton.onclick = (): void => {
      this._isCancelled = true;
      cancelButton.disabled = true;
      status.textContent = strings.CancellingMessage;
      status.className = `${styles.status}`;
    };
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

  /**
   * Read keyword text from SharePoint / Fluent / PnP SearchBox-style fields (best-scoring visible input).
   */
  private _tryReadSharePointSearchBoxValue(): string {
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
    const hosts = document.querySelectorAll(
      'div[role="combobox"], button[aria-haspopup="listbox"], div[aria-expanded]'
    );
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

    const labels = document.querySelectorAll('label, span, div');
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

    return undefined;
  }

  private _escapeFqlEqualsArg(value: string): string {
    return (value || '')
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"');
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

  /**
   * PnP date refiners belong in RefinementFilters (FQL), not KQL Querytext — same path as FileType in the refiners web part.
   * Uses RANGE / equals patterns from MS FQL reference.
   */
  private _buildDateRefinerFqlGroup(group: IPnpFilterGroup, managedProperty: string): string {
    const vals = group.values || [];
    type Dated = { op: number; iso: string };
    const dated: Dated[] = [];
    for (let i = 0; i < vals.length; i++) {
      const raw = (vals[i].value || '').trim();
      if (!raw || !this._looksLikeIsoDateTimeForKql(raw)) {
        continue;
      }
      const op =
        vals[i].operator !== undefined && vals[i].operator !== null
          ? Number(vals[i].operator)
          : PnpFilterComparisonOperator.Eq;
      dated.push({ op, iso: raw });
    }
    if (dated.length === 0) {
      return '';
    }

    const innerJoin = String(group.operator || 'or').toLowerCase() === 'and' ? 'and' : 'or';

    if (dated.length === 2 && innerJoin === 'and') {
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
        return `${managedProperty}:range(${this._fqlDatetimeOperandFromIso(geD.iso)}, ${this._fqlDatetimeOperandFromIso(
          leD.iso
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
    return innerJoin === 'and' ? `and(${parts.join(', ')})` : `or(${parts.join(', ')})`;
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
    const innerOp = innerOpLower === 'and' ? ' AND ' : ' OR ';
    return `(${pieces.join(innerOp)})`;
  }

  /**
   * Map PnP URL `f` JSON to FQL (RefinementFilters) for refiners that behave like the Filters web part,
   * and KQL only for remaining refiner types. FileType + date refiners use FQL so totals align with Search Results.
   */
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

      if (
        fn === 'created' ||
        fn === 'lastmodifiedtime' ||
        fn === 'lastmodified' ||
        fn === 'modified'
      ) {
        let managed = prop;
        if (fn === 'lastmodifiedtime' || fn === 'lastmodified' || fn === 'modified') {
          managed = 'LastModifiedTime';
        } else {
          managed = 'Created';
        }
        const dg = this._buildDateRefinerFqlGroup(group, managed);
        if (dg) {
          fqlGroupClauses.push(dg);
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
    // Default: do NOT apply URL refiners unless explicitly enabled.
    // (If we treat `undefined` as enabled, leftover `f=` can easily filter everything out.)
    const useFilters = this.properties.appendUrlFilters === true;
    if (!useFilters) {
      return { filterKql: '', refinementFql: '', summary: strings.FiltersDisabledLabel };
    }

    const paramKey = 'f';
    let raw: string | undefined;
    let summaryPrefix: string | undefined;

    const named = this._getUrlParam(paramKey);
    if (named && named.trim()) {
      raw = named.trim();
      summaryPrefix = `${strings.FromUrlLabel} (${paramKey})`;
    } else {
      const discovered = this._discoverPnpRefinersInUrlParams();
      if (discovered) {
        raw = discovered.raw.trim();
        summaryPrefix = `${strings.FiltersDiscoveredInUrlLabel} (${discovered.fromKey})`;
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

  private _getCellValue(cells: Array<{ Key?: unknown; Value?: unknown }>, candidates: string[]): string | undefined {
    const lowered = candidates.map((c) => c.toLowerCase());
    for (let i = 0; i < cells.length; i++) {
      const key = this._normalizeToString(cells[i]?.Key).toLowerCase();
      for (let j = 0; j < lowered.length; j++) {
        if (key.indexOf(lowered[j]) !== -1) {
          return this._normalizeToString(cells[i]?.Value);
        }
      }
    }
    return undefined;
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

  private _extractRowsFromSearchJson(json: unknown): {
    rows: Array<{ Title: string; Path: string; Author: string }>;
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

    const mapped = rows.map((row) => {
      const rowObj = row as { Cells?: unknown };
      const cells = toCellsArray(rowObj?.Cells);
      return {
        Title: this._normalizeToString(this._getCellValue(cells, ['Title']) || ''),
        Path: this._normalizeToString(this._getCellValue(cells, ['Path']) || ''),
        Author: this._normalizeToString(this._getCellValue(cells, ['Author']) || '')
      };
    });

    const lastRow = rows.length > 0 ? rows[rows.length - 1] : undefined;
    const lastRowObj = lastRow as { Cells?: unknown } | undefined;
    const lastCells = toCellsArray(lastRowObj?.Cells);
    const lastDocIdRaw = this._getCellValue(lastCells, ['IndexDocId', 'indexdocid', 'DocId', 'docid']);
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
    /** FQL refinement strings (e.g. `FileType:equals("docx")`) — same mechanism as PnP filter web part. */
    refinementFilters?: string[];
  }): Promise<{
    rows: Array<{ Title: string; Path: string; Author: string }>;
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
    const selectProps = params.selectProperties
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
      TrimDuplicates: false
    };
    if (refinementList.length > 0) {
      requestBody.RefinementFilters = refinementList;
    }
    const payload = { request: requestBody };

    // Response JSON varies (`d.postquery` vs `postquery`, verbose vs nometadata). Try verbose first, then minimal.
    let json: unknown;
    let odataAttempt = 'verbose';
    try {
      json = await this._postSearchPostquery(postUrl, payload, 'verbose');
    } catch (firstErr) {
      try {
        odataAttempt = 'nometadata(fallback)';
        json = await this._postSearchPostquery(postUrl, payload, 'nometadata');
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
        const json2 = await this._postSearchPostquery(postUrl, payload, 'nometadata');
        const found2 = this._findRelevantResultsBlock(json2);
        if (found2.block && (rowCountFromBlock(found2.block) > 0 || !found.block)) {
          json = json2;
          found = found2;
          odataAttempt =
            odataAttempt.indexOf('fallback') !== -1 ? 'verbose+nometadata' : 'verbose-then-nometadata';
        }
      } catch {
        // keep first json
      }
    }

    let transport = `postquery:${odataAttempt}`;
    let ex = this._extractRowsFromSearchJson(json);

    if (ex.rows.length === 0) {
      try {
        const jsonGet = await this._getSearchQueryViaGet(
          webUrl,
          safeQuery,
          sourceId,
          params.pageSize,
          selectProps,
          refinementFqlForGet
        );
        const exGet = this._extractRowsFromSearchJson(jsonGet);
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('sourceId', {
                  label: strings.SourceIdLabel,
                  description: strings.SourceIdDescription
                }),
                PropertyPaneToggle('appendUrlFilters', {
                  label: strings.AppendUrlFiltersLabel
                }),
                PropertyPaneToggle('debugApi', {
                  label: strings.DebugApiLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
