import { PropertyPaneTextField } from '@microsoft/sp-property-pane';

/** Strings required for the Export button property-pane group (extend when adding fields). */
export interface IExportButtonPaneStrings {
  ButtonAppearanceGroupLabel: string;
  ExportButtonTextLabel: string;
  ExportButtonTextDescription: string;
  CancelButtonTextLabel: string;
  CancelButtonTextDescription: string;
  ExportButtonBackgroundLabel: string;
  ExportButtonTextColorLabel: string;
  ExportButtonBorderColorLabel: string;
  ExportButtonBorderRadiusLabel: string;
  ExportButtonFontSizeLabel: string;
  ExportButtonWidthLabel: string;
  ExportButtonHeightLabel: string;
  ExportButtonColorFieldsDescription: string;
  ExportButtonRadiusDescription: string;
  ExportButtonLengthFieldsDescription: string;
}

/** Return type matches `groupFields` in `IPropertyPaneGroup` (SPFx generic varies by field type). */
export function buildExportButtonAppearanceGroupFields(strings: IExportButtonPaneStrings): ReturnType<typeof PropertyPaneTextField>[] {
  return [
    PropertyPaneTextField('exportButtonText', {
      label: strings.ExportButtonTextLabel,
      description: strings.ExportButtonTextDescription
    }),
    PropertyPaneTextField('cancelButtonText', {
      label: strings.CancelButtonTextLabel,
      description: strings.CancelButtonTextDescription
    }),
    PropertyPaneTextField('exportButtonBackgroundColor', {
      label: strings.ExportButtonBackgroundLabel,
      description: strings.ExportButtonColorFieldsDescription
    }),
    PropertyPaneTextField('exportButtonTextColor', {
      label: strings.ExportButtonTextColorLabel,
      description: strings.ExportButtonColorFieldsDescription
    }),
    PropertyPaneTextField('exportButtonBorderColor', {
      label: strings.ExportButtonBorderColorLabel,
      description: strings.ExportButtonColorFieldsDescription
    }),
    PropertyPaneTextField('exportButtonBorderRadius', {
      label: strings.ExportButtonBorderRadiusLabel,
      description: strings.ExportButtonRadiusDescription
    }),
    PropertyPaneTextField('exportButtonFontSize', {
      label: strings.ExportButtonFontSizeLabel,
      description: strings.ExportButtonLengthFieldsDescription
    }),
    PropertyPaneTextField('exportButtonWidth', {
      label: strings.ExportButtonWidthLabel,
      description: strings.ExportButtonLengthFieldsDescription
    }),
    PropertyPaneTextField('exportButtonHeight', {
      label: strings.ExportButtonHeightLabel,
      description: strings.ExportButtonLengthFieldsDescription
    })
  ];
}
