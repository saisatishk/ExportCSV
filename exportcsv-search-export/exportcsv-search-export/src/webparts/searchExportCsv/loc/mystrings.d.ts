declare interface ISearchExportCsvWebPartStrings {
  ResolvedQueryLabel: string;
  SourceIdLabel: string;
  ExportColumnsLabel: string;
  ExportColumnsDescription: string;
  DebugApiLabel: string;
  WebPartTitle: string;
  WebPartDescription: string;
  QuerySourceLabel: string;
  SourceIdSourceLabel: string;
  FromPropertyLabel: string;
  FromUrlLabel: string;
  NotSetLabel: string;
  ExportButtonLabel: string;
  CancelButtonLabel: string;
  SourceIdRequiredError: string;
  InvalidSourceIdGuidError: string;
  ExportStarted: string;
  ExportInProgress: string;
  PageLabel: string;
  ExportCancelled: string;
  CancellingMessage: string;
  ExportCappedMessage: string;
  ExportCompleted: string;
  ExportFailedPrefix: string;
  FiltersFromUrlLabel: string;
  NoFiltersInUrlLabel: string;
  FiltersParseFailedLabel: string;
  FiltersDiscoveredInUrlLabel: string;
  FiltersFromUiLabel: string;
  EffectiveFilterKqlLabel: string;
  SearchQueryFromPageSearchBoxLabel: string;
  ExportNoKeywordsWithRefinersHint: string;
  ExportNoKeywordsNoRefinersHint: string;
}

declare module 'SearchExportCsvWebPartStrings' {
  const strings: ISearchExportCsvWebPartStrings;
  export = strings;
}
