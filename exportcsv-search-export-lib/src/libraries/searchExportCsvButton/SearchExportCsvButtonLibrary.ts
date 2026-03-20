import {
  IComponentDefinition,
  IExtensibilityLibrary,
  IDataSourceDefinition,
  ILayoutDefinition,
  IAdaptiveCardAction,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition
} from '@pnp/modern-search-extensibility';

import { SearchExportCsvButtonWebComponent, ensureSearchExportCsvButtonRegistered } from './SearchExportCsvButtonWebComponent';

// Ensure the custom element is defined as soon as this library bundle loads.
// This helps when the extensibility hook isn't invoked immediately.
ensureSearchExportCsvButtonRegistered();

/**
 * PnP Modern Search extensibility library entry point.
 * Registers the custom web component used to render the Export CSV button in search results templates.
 */
export class SearchExportCsvButtonLibrary implements IExtensibilityLibrary {
  public name(): string {
    return 'SearchExportCsvButtonLibrary';
  }

  public getCustomWebComponents(): IComponentDefinition<typeof SearchExportCsvButtonWebComponent>[] {
    ensureSearchExportCsvButtonRegistered();

    return [
      {
        componentName: 'search-export-csv-button',
        componentClass: SearchExportCsvButtonWebComponent
      }
    ];
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  public getCustomQueryModifiers?(): IQueryModifierDefinition[] {
    return [];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [];
  }

  public invokeCardAction(_action: IAdaptiveCardAction): void {
    // This extension doesn't render adaptive card actions.
  }
}
