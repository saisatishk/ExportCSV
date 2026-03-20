declare interface ISearchExportCsvWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ResolvedQueryLabel: string;
  SourceIdLabel: string;
  SourceIdDescription: string;
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
  FiltersDisabledLabel: string;
  NoFiltersInUrlLabel: string;
  FiltersParseFailedLabel: string;
  AppendUrlFiltersLabel: string;
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
