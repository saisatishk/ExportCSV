/** Label from property or localized default. */
export function resolveButtonLabel(raw: string | undefined, fallback: string): string {
  const t = (raw || '').trim();
  return t.length > 0 ? t : fallback;
}

export function sanitizeCssColor(raw: string | undefined): string | undefined {
  const s = (raw || '').trim();
  if (!s) return undefined;
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(s) ? s : undefined;
}

export function sanitizeCssLength(raw: string | undefined): string | undefined {
  const s = (raw || '').trim();
  if (!s) return undefined;
  return /^\d+(\.\d+)?(px|rem|em|%)?$/.test(s) ? s : undefined;
}

export interface IExportButtonLabelProps {
  exportButtonText?: string;
  cancelButtonText?: string;
}

export interface IExportButtonStyleProps {
  exportButtonBackgroundColor?: string;
  exportButtonTextColor?: string;
  exportButtonBorderColor?: string;
  exportButtonBorderRadius?: string;
  exportButtonFontSize?: string;
  exportButtonWidth?: string;
  exportButtonHeight?: string;
}

export type IExportButtonAppearanceProps = IExportButtonLabelProps & IExportButtonStyleProps;

/** Safe ` style="..."` fragment from property pane (hex colors / length only). */
export function buildExportButtonStyleAttr(style: IExportButtonStyleProps): string {
  const parts: string[] = [];
  const bg = sanitizeCssColor(style.exportButtonBackgroundColor);
  const fg = sanitizeCssColor(style.exportButtonTextColor);
  const bc = sanitizeCssColor(style.exportButtonBorderColor);
  const br = sanitizeCssLength(style.exportButtonBorderRadius);
  const fs = sanitizeCssLength(style.exportButtonFontSize);
  const w = sanitizeCssLength(style.exportButtonWidth);
  const h = sanitizeCssLength(style.exportButtonHeight);
  if (bg) parts.push(`background:${bg}`);
  if (fg) parts.push(`color:${fg}`);
  if (bc) parts.push(`border-color:${bc}`);
  if (br) parts.push(`border-radius:${br}`);
  if (fs) parts.push(`font-size:${fs}`);
  if (w) parts.push(`width:${w}`);
  if (h) parts.push(`height:${h}`);
  return parts.length > 0 ? ` style="${parts.join(';')}"` : '';
}
