import { getColorFromString } from '@fluentui/react/lib/Color';
import { ComponentsStyles, createTheme, IFontStyles, IPalette, ISemanticColors, ITheme } from '@fluentui/react/lib/Theme';
import { safeColor } from './sanitize';
import { asRecord, IUntyped } from './untyped';

// SPFx 1.4.1 has no theme service: site colors come from window.__themeState__.theme, createTheme derives the rest.
const COMPONENT_OVERRIDES: ComponentsStyles = {
  // Links must not be recognizable by their color alone (WCAG 1.4.1).
  Link: {
    styles: {
      root: { textDecoration: 'underline' }
    }
  }
};

/** Accessibility requirement. */
const MIN_FONT_SIZE = '14px';

/** WCAG 2.1 AA: text (1.4.3) and graphics such as icons (1.4.11). */
const MIN_TEXT_CONTRAST = 4.5;
export const MIN_GRAPHIC_CONTRAST = 3;

/** Disabled elements are left out (exempt in WCAG). */
const TEXT_PAIRS: Array<[keyof ISemanticColors, keyof ISemanticColors]> = [
  ['bodyText', 'bodyBackground'],
  ['bodySubtext', 'bodyBackground'],
  ['bodySubtext', 'bodyStandoutBackground'],
  ['errorText', 'bodyBackground'],
  ['link', 'bodyBackground'],
  ['linkHovered', 'bodyBackground'],
  ['actionLink', 'bodyBackground'],
  ['inputText', 'inputBackground'],
  ['inputPlaceholderText', 'inputBackground'],
  ['buttonText', 'buttonBackground'],
  ['buttonTextHovered', 'buttonBackgroundHovered'],
  ['buttonTextPressed', 'buttonBackgroundPressed'],
  ['primaryButtonText', 'primaryButtonBackground'],
  ['primaryButtonTextHovered', 'primaryButtonBackgroundHovered'],
  ['primaryButtonTextPressed', 'primaryButtonBackgroundPressed'],
  ['menuItemText', 'menuBackground'],
  ['messageText', 'infoBackground'],
  ['messageText', 'errorBackground'],
  ['messageText', 'blockingBackground'],
  ['messageText', 'warningBackground'],
  ['messageText', 'severeWarningBackground'],
  ['messageText', 'successBackground'],
  ['warningText', 'warningBackground']
];

const ICON_PAIRS: Array<[keyof ISemanticColors, keyof ISemanticColors]> = [
  ['successIcon', 'bodyBackground'],
  ['warningIcon', 'bodyBackground'],
  ['errorIcon', 'bodyBackground'],
  ['inputBorder', 'inputBackground'],
  ['inputBorderHovered', 'inputBackground'],
  ['inputFocusBorderAlt', 'inputBackground']
];

export function createSharePointTheme(): ITheme {
  const defaults = createTheme({ components: COMPONENT_OVERRIDES });
  const slots = asRecord(asRecord(asRecord(window).__themeState__).theme);

  const palette = pickColors(slots, Object.keys(defaults.palette));
  const semanticColors = Object.keys(palette).length > 0 ? pickColors(slots, Object.keys(defaults.semanticColors)) : {};
  const build = (semantic: { [slot: string]: string }): ITheme => createTheme({
    palette: palette as Partial<IPalette>,
    semanticColors: semantic as Partial<ISemanticColors>,
    fonts: minimumFontSizes(defaults),
    components: COMPONENT_OVERRIDES,
    isInverted: isDark(palette.white)
  });

  // Second pass only if the site colors break a contrast: the fixes are semantic slots of their own.
  const theme = build(semanticColors);
  const fixes = contrastFixes(theme);
  return Object.keys(fixes).length > 0 ? build({ ...semanticColors, ...fixes }) : theme;
}

function minimumFontSizes(theme: ITheme): Partial<IFontStyles> {
  const fonts: Partial<IFontStyles> = {};
  (['tiny', 'xSmall', 'small', 'smallPlus'] as Array<keyof IFontStyles>).forEach((name: keyof IFontStyles) => {
    fonts[name] = { ...theme.fonts[name], fontSize: MIN_FONT_SIZE };
  });
  return fonts;
}

function contrastFixes(theme: ITheme): { [slot: string]: string } {
  const colors = theme.semanticColors;
  const fixes: { [slot: string]: string } = {};
  const check = (pairs: Array<[keyof ISemanticColors, keyof ISemanticColors]>, minimum: number): void => {
    pairs.forEach(([text, background]: [keyof ISemanticColors, keyof ISemanticColors]) => {
      const current = fixes[text] || colors[text];
      const readable = readableColor(colors[background], [current, colors.bodyText], minimum);
      if (readable !== current) {
        fixes[text] = readable;
      }
    });
  };
  check(TEXT_PAIRS, MIN_TEXT_CONTRAST);
  check(ICON_PAIRS, MIN_GRAPHIC_CONTRAST);
  return fixes;
}

/** Unparseable colors are kept. */
export function readableColor(background: string, candidates: string[], minimum: number = MIN_TEXT_CONTRAST): string {
  if (!getColorFromString(background)) {
    return candidates[0];
  }
  for (const candidate of candidates) {
    const ratio = contrastRatio(candidate, background);
    if (ratio === undefined || ratio >= minimum) {
      return candidate;
    }
  }
  return (contrastRatio('#000000', background) || 0) >= (contrastRatio('#ffffff', background) || 0) ? '#000000' : '#ffffff';
}

function contrastRatio(a: string, b: string): number | undefined {
  const la = luminance(a);
  const lb = luminance(b);
  if (la === undefined || lb === undefined) {
    return undefined;
  }
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/** Relative luminance (WCAG 2.1). */
function luminance(value: string): number | undefined {
  const color = value ? getColorFromString(value) : undefined;
  if (!color) {
    return undefined;
  }
  const channel = (c: number): number => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(color.r) + 0.7152 * channel(color.g) + 0.0722 * channel(color.b);
}

/** The slot values are untrusted: only valid CSS colors. */
function pickColors(slots: IUntyped, names: string[]): { [slot: string]: string } {
  const colors: { [slot: string]: string } = {};
  names.forEach((name: string) => {
    const color = safeColor(slots[name], '');
    if (color) { colors[name] = color; }
  });
  return colors;
}

/** In dark themes the "white" slot (page background) is dark. */
function isDark(background: string | undefined): boolean {
  const color = background ? getColorFromString(background) : undefined;
  return !!color && (0.299 * color.r + 0.587 * color.g + 0.114 * color.b) < 128;
}
