import { IButtonStyles } from '@fluentui/react/lib/Button';
import { FontWeights, HighContrastSelector, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ITheme, useTheme } from '@fluentui/react/lib/Theme';
import { memoizeFunction } from '@fluentui/react/lib/Utilities';
import { MIN_GRAPHIC_CONTRAST, readableColor } from '../utils/theme';

const GUTTER = 16;

export interface IFeedbackClassNames {
  columns: string;
  column: string;
  sectionTitle: string;
  field: string;
  hint: string;
  message: string;
}

const getClassNames = memoizeFunction((theme: ITheme): IFeedbackClassNames => mergeStyleSets({
  columns: {
    display: 'flex',
    flexWrap: 'wrap',
    margin: `0 -${GUTTER}px`
  },
  column: {
    flex: '1 1 320px',
    minWidth: 0,
    boxSizing: 'border-box',
    padding: `0 ${GUTTER}px`,
    border: 0,
    margin: 0
  },
  sectionTitle: {
    ...theme.fonts.mediumPlus,
    fontWeight: FontWeights.semibold,
    color: theme.semanticColors.bodyText,
    margin: '8px 0 4px',
    padding: 0
  },
  field: {
    marginTop: 12
  },
  hint: {
    ...theme.fonts.medium,
    color: theme.semanticColors.bodySubtext,
    margin: '5px 0 0'
  },
  message: {
    marginBottom: 12
  }
}));

export function useClassNames(): IFeedbackClassNames {
  return getClassNames(useTheme());
}

export interface ISettingsClassNames {
  tabs: string;
  panel: string;
  notice: string;
  description: string;
  warning: string;
  links: string;
  row: string;
  label: string;
  value: string;
  problem: string;
  ok: string;
  language: string;
  languageNote: string;
  intro: string;
  checklist: string;
  checkItem: string;
  checkIcon: string;
  iconOk: string;
  iconWarning: string;
  iconError: string;
  checkLabel: string;
  checkDetail: string;
  checkProblem: string;
  actions: string;
  loading: string;
}

const getSettingsClassNames = memoizeFunction((theme: ITheme): ISettingsClassNames => {
  const text = { ...theme.fonts.medium, margin: 0, overflowWrap: 'break-word' as const, wordBreak: 'break-word' as const };
  const small = { ...theme.fonts.medium, margin: '5px 0 0', overflowWrap: 'break-word' as const, wordBreak: 'break-word' as const };
  const background = theme.semanticColors.bodyBackground;
  return mergeStyleSets({
    tabs: {
      marginBottom: 4
    },
    panel: {
      minHeight: 340
    },
    notice: {
      marginTop: 12
    },
    description: { ...small, color: theme.semanticColors.bodySubtext },
    warning: { ...small, color: readableColor(background, [theme.semanticColors.warningText, theme.semanticColors.bodyText]) },
    links: {
      ...theme.fonts.medium,
      margin: '6px 0 0',
      selectors: { '& > *': { marginRight: 16 } }
    },
    row: {
      marginTop: 12
    },
    label: {
      ...theme.fonts.medium,
      fontWeight: FontWeights.semibold,
      color: theme.semanticColors.bodyText,
      padding: '5px 0'
    },
    value: { ...text, color: theme.semanticColors.bodyText },
    problem: { ...text, color: theme.semanticColors.errorText, marginTop: 12 },
    ok: { ...text, color: readableColor(background, [theme.semanticColors.successIcon, theme.semanticColors.bodyText]), marginTop: 12 },
    language: {
      marginTop: 12,
      maxWidth: 320
    },
    languageNote: { ...small, color: theme.semanticColors.bodySubtext, marginBottom: 8 },
    intro: { ...text, color: theme.semanticColors.bodyText, marginTop: 12 },
    checklist: {
      listStyle: 'none',
      margin: '12px 0 0',
      padding: 0
    },
    checkItem: {
      display: 'flex',
      alignItems: 'flex-start',
      padding: '10px 0',
      borderBottom: `1px solid ${theme.semanticColors.bodyDivider}`
    },
    checkIcon: {
      flexShrink: 0,
      fontSize: 16,
      lineHeight: 20,
      marginRight: 12
    },
    iconOk: { color: theme.semanticColors.successIcon },
    iconWarning: { color: theme.semanticColors.warningIcon },
    iconError: { color: theme.semanticColors.errorIcon },
    checkLabel: { ...text, color: theme.semanticColors.bodyText },
    checkDetail: { ...small, margin: '2px 0 0', color: theme.semanticColors.bodySubtext },
    checkProblem: { ...small, margin: '2px 0 0', color: theme.semanticColors.errorText },
    actions: {
      marginTop: 16,
      selectors: { '& > *': { marginRight: 8, marginBottom: 8 } }
    },
    loading: {
      minHeight: 340
    }
  });
});

export function useSettingsClassNames(): ISettingsClassNames {
  return getSettingsClassNames(useTheme());
}

export interface IPreviewClassNames {
  title: string;
  hint: string;
  mail: string;
  header: string;
  label: string;
  value: string;
  subject: string;
  address: string;
  none: string;
  body: string;
}

const getPreviewClassNames = memoizeFunction((theme: ITheme): IPreviewClassNames => {
  const wrap = { overflowWrap: 'break-word' as const, wordBreak: 'break-word' as const };
  const cell = { padding: '6px 0', verticalAlign: 'top', borderBottom: `1px solid ${theme.semanticColors.bodyDivider}` };
  return mergeStyleSets({
    title: {
      ...theme.fonts.mediumPlus,
      fontWeight: FontWeights.semibold,
      color: theme.semanticColors.bodyText,
      margin: '8px 0 4px'
    },
    hint: {
      ...theme.fonts.medium,
      color: theme.semanticColors.bodySubtext,
      margin: '0 0 12px'
    },
    mail: {
      border: `1px solid ${theme.semanticColors.bodyDivider}`,
      borderRadius: theme.effects.roundedCorner4,
      overflow: 'hidden',
      boxShadow: theme.effects.elevation4
    },
    header: {
      width: '100%',
      borderCollapse: 'collapse',
      tableLayout: 'auto',
      backgroundColor: theme.semanticColors.bodyStandoutBackground,
      selectors: { 'tr:last-child th, tr:last-child td': { borderBottom: 0 } }
    },
    label: {
      ...theme.fonts.medium,
      ...cell,
      width: '1%',
      whiteSpace: 'nowrap',
      textAlign: 'left',
      fontWeight: FontWeights.semibold,
      color: theme.semanticColors.bodySubtext,
      padding: '8px 24px 6px 16px'
    },
    value: {
      ...theme.fonts.medium,
      ...wrap,
      ...cell,
      color: theme.semanticColors.bodyText,
      padding: '4px 16px 4px 0'
    },
    subject: {
      ...theme.fonts.large,
      ...wrap,
      ...cell,
      fontWeight: FontWeights.semibold,
      color: theme.semanticColors.bodyText,
      padding: '6px 16px 6px 0'
    },
    address: {
      display: 'inline-block',
      maxWidth: '100%',
      boxSizing: 'border-box',
      ...wrap,
      margin: '2px 6px 2px 0',
      padding: '1px 10px',
      borderRadius: 12,
      backgroundColor: theme.palette.themeLighterAlt,
      border: `1px solid ${theme.palette.themeLighter}`,
      color: readableColor(theme.palette.themeLighterAlt, [theme.palette.themeDarker, theme.semanticColors.bodyText]),
      selectors: { [HighContrastSelector]: { border: '1px solid WindowText', color: 'WindowText' } }
    },
    none: {
      display: 'inline-block',
      padding: '3px 0',
      fontStyle: 'italic',
      color: theme.semanticColors.bodySubtext
    },
    body: {
      ...wrap,
      padding: 16,
      maxHeight: '50vh',
      overflowY: 'auto',
      backgroundColor: '#ffffff',
      borderTop: `1px solid ${theme.semanticColors.bodyDivider}`,
      selectors: {
        a: { color: '#0078d4' },
        [HighContrastSelector]: { backgroundColor: 'Window', color: 'WindowText' }
      }
    }
  });
});

export function usePreviewClassNames(): IPreviewClassNames {
  return getPreviewClassNames(useTheme());
}

export const getFloatingButtonStyles = memoizeFunction((theme: ITheme): IButtonStyles => {
  const { palette } = theme;
  const iconOn = (background: string): string => readableColor(background, [palette.white, palette.neutralPrimary], MIN_GRAPHIC_CONTRAST);
  const edge = readableColor(palette.white, [palette.themePrimary, palette.neutralPrimary], MIN_GRAPHIC_CONTRAST);
  return {
    root: {
      position: 'fixed',
      right: 24,
      bottom: 24,
      zIndex: 1000,
      width: 52,
      height: 52,
      borderRadius: '50%',
      border: `2px solid ${edge === palette.themePrimary ? 'transparent' : edge}`,
      backgroundColor: palette.themePrimary,
      color: iconOn(palette.themePrimary),
      boxShadow: theme.effects.elevation16,
      selectors: {
        [HighContrastSelector]: { borderColor: 'ButtonText', backgroundColor: 'ButtonFace', color: 'ButtonText' },
        '@media print': { display: 'none' }
      }
    },
    rootHovered: { backgroundColor: palette.themeDarkAlt, color: iconOn(palette.themeDarkAlt) },
    rootPressed: { backgroundColor: palette.themeDark, color: iconOn(palette.themeDark) },
    rootFocused: { outline: `2px solid ${palette.neutralPrimary}`, outlineOffset: '2px' },
    icon: { fontSize: 22, margin: 0, color: 'inherit' },
    iconHovered: { color: 'inherit' },
    iconPressed: { color: 'inherit' }
  };
});
