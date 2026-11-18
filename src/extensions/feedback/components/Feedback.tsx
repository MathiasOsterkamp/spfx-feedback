import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { ITheme, ThemeProvider } from '@fluentui/react/lib/Theme';
import { FeedbackDialog } from './FeedbackDialog';
import { getFloatingButtonStyles } from './styles';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackOrigin, IPageInfo } from '../models/IFeedback';
import { MailService } from '../services/MailService';
import { RecipientService } from '../services/RecipientService';
import { SettingsService } from '../services/SettingsService';
import { SetupService } from '../services/SetupService';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface IFeedbackProps {
  /** Theme of the SharePoint site, calculated by the extension (see utils/theme.ts) */
  theme: ITheme;
  /** Current page; undefined if it is not an item of the Site Pages library -> no button. */
  page: IPageInfo | undefined;
  origin: IFeedbackOrigin;
  policy: ISecurityPolicy;
  /** Code configuration and settings list of the site (read when the dialog opens). */
  settingsService: SettingsService;
  mailService: MailService;
  recipientService: RecipientService;
  setupService: SetupService;
  /** Current user may open the settings (gear icon): at least "Design" permission on the site. */
  canConfigure: boolean;
  /** Current user may store entries in the settings list (once it is set up). */
  canSaveSettings: boolean;
}

/**
 * Address of the current page without query string and fragment: removes parameters that do not
 * belong to the page (e.g. the debug parameters loadSPFX/debugManifestsFile/customActions).
 */
function currentPageUrl(): string {
  const location = window.location;
  return location.protocol + '//' + location.host + location.pathname;
}

/** Floating feedback button; opens the dialog "Feedback geben / Inhalte melden". */
export const Feedback: React.FunctionComponent<IFeedbackProps> = (props: IFeedbackProps) => {
  const { theme, page, ...dialogProps } = props;
  const [openedOn, setOpenedOn] = React.useState<{ page: IPageInfo; url: string } | undefined>(undefined);

  if (!page) {
    return null;
  }

  // Modern pages navigate without reloading: page and address are taken at the moment of opening.
  const open = (): void => setOpenedOn({ page, url: currentPageUrl() });

  // applyTo="none": only provides the theme, without painting a background over the page.
  return (
    <ThemeProvider theme={theme} applyTo="none">
      <IconButton
        iconProps={{ iconName: 'Feedback' }}
        title={strings.buttonLabel}
        ariaLabel={strings.buttonLabel}
        aria-haspopup="dialog"
        styles={getFloatingButtonStyles(theme)}
        onClick={open}
      />
      {openedOn && (
        <FeedbackDialog
          {...dialogProps}
          page={openedOn.page}
          pageUrl={openedOn.url}
          onDismiss={() => setOpenedOn(undefined)}
        />
      )}
    </ThemeProvider>
  );
};
