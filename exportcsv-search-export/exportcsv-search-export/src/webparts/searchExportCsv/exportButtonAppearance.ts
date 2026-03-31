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

export interface IExportButtonStyleProps {
  exportButtonBackgroundColor?: string;
  exportButtonTextColor?: string;
  exportButtonBorderColor?: string;
  exportButtonBorderRadius?: string;
}

/** Safe ` style="..."` fragment from property pane (hex colors / length only). */
export function buildExportButtonStyleAttr(style: IExportButtonStyleProps): string {
  const parts: string[] = [];
  const bg = sanitizeCssColor(style.exportButtonBackgroundColor);
  const fg = sanitizeCssColor(style.exportButtonTextColor);
  const bc = sanitizeCssColor(style.exportButtonBorderColor);
  const br = sanitizeCssLength(style.exportButtonBorderRadius);
  if (bg) parts.push(`background:${bg}`);
  if (fg) parts.push(`color:${fg}`);
  if (bc) parts.push(`border-color:${bc}`);
  if (br) parts.push(`border-radius:${br}`);
  return parts.length > 0 ? ` style="${parts.join(';')}"` : '';
}
