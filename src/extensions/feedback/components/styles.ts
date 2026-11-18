import { IButtonStyles } from '@fluentui/react/lib/Button';
import { FontWeights, HighContrastSelector, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ITheme, useTheme } from '@fluentui/react/lib/Theme';
import { memoizeFunction } from '@fluentui/react/lib/Utilities';

/** Gap between the two columns and around the form. */
const GUTTER = 16;

/** Generated class names of the dialog (explicit type: needed for the declaration output of the build). */
export interface IFeedbackClassNames {
  columns: string;
  column: string;
  sectionTitle: string;
  hint: string;
  field: string;
  message: string;
}

/** Styles depend on the theme, therefore they are created per theme (and cached). */
const getClassNames = memoizeFunction((theme: ITheme): IFeedbackClassNames => mergeStyleSets({
  // Flexbox with negative margins instead of "gap": also works in IE11 (SharePoint 2019).
  columns: {
    display: 'flex',
    flexWrap: 'wrap',
    margin: `0 -${GUTTER}px`
  },
  column: {
    // Two columns from about 2 × 320 px, below that one under the other.
    flex: '1 1 320px',
    minWidth: 0,
    boxSizing: 'border-box',
    padding: `0 ${GUTTER}px`,
    // Reset of the browser styles for <fieldset>
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
  hint: {
    ...theme.fonts.small,
    color: theme.semanticColors.bodySubtext,
    margin: '8px 0 0'
  },
  field: {
    marginTop: 12
  },
  message: {
    marginBottom: 12
  }
}));

export function useClassNames(): IFeedbackClassNames {
  return getClassNames(useTheme());
}

/** Class names of the setup view (SetupView). */
export interface ISetupClassNames {
  row: string;
  label: string;
  value: string;
  ok: string;
  problem: string;
  action: string;
  hint: string;
  visibility: string;
  loading: string;
}

const getSetupClassNames = memoizeFunction((theme: ITheme): ISetupClassNames => {
  const text = { ...theme.fonts.medium, margin: 0, overflowWrap: 'break-word' as const, wordBreak: 'break-word' as const };
  return mergeStyleSets({
    // Read-only "label / value" rows, aligned like the labels of the text fields.
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
    ok: { ...text, color: theme.semanticColors.successIcon },
    problem: { ...text, color: theme.semanticColors.errorText },
    action: {
      marginTop: 16
    },
    hint: {
      ...theme.fonts.small,
      color: theme.semanticColors.bodySubtext,
      margin: '0 0 8px'
    },
    visibility: {
      ...theme.fonts.small,
      color: theme.semanticColors.bodySubtext,
      margin: '24px 0 0'
    },
    loading: {
      minHeight: 160
    }
  });
});

export function useSetupClassNames(): ISetupClassNames {
  return getSetupClassNames(useTheme());
}

/** Round, floating button at the bottom right of the page. */
export const getFloatingButtonStyles = memoizeFunction((theme: ITheme): IButtonStyles => ({
  root: {
    position: 'fixed',
    right: 24,
    bottom: 24,
    // Above the page content, below Fluent layers (dialogs, panels: 1000000).
    zIndex: 1000,
    width: 52,
    height: 52,
    borderRadius: '50%',
    border: '2px solid transparent',
    backgroundColor: theme.palette.themePrimary,
    color: theme.palette.white,
    boxShadow: theme.effects.elevation16,
    selectors: {
      [HighContrastSelector]: { borderColor: 'ButtonText', backgroundColor: 'ButtonFace', color: 'ButtonText' },
      '@media print': { display: 'none' }
    }
  },
  rootHovered: { backgroundColor: theme.palette.themeDarkAlt, color: theme.palette.white },
  rootPressed: { backgroundColor: theme.palette.themeDark, color: theme.palette.white },
  rootFocused: { outline: `2px solid ${theme.palette.neutralPrimary}`, outlineOffset: '2px' },
  icon: { fontSize: 22, margin: 0, color: 'inherit' },
  iconHovered: { color: 'inherit' },
  iconPressed: { color: 'inherit' }
}));
