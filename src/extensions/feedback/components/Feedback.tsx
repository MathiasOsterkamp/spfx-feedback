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
  theme: ITheme;
  page: IPageInfo | undefined;
  origin: IFeedbackOrigin;
  policy: ISecurityPolicy;
  settingsService: SettingsService;
  mailService: MailService;
  recipientService: RecipientService;
  setupService: SetupService;
  canConfigure: boolean;
  canSaveSettings: boolean;
}

/** Without query and fragment: drops parameters such as loadSPFX/debugManifestsFile. */
function currentPageUrl(): string {
  const location = window.location;
  return location.protocol + '//' + location.host + location.pathname;
}

export const Feedback: React.FunctionComponent<IFeedbackProps> = (props: IFeedbackProps) => {
  const { theme, page, ...dialogProps } = props;
  const [openedOn, setOpenedOn] = React.useState<{ page: IPageInfo; url: string } | undefined>(undefined);

  if (!page) {
    return null;
  }

  // Modern pages navigate without reloading: take page and address when opening.
  const open = (): void => setOpenedOn({ page, url: currentPageUrl() });

  // applyTo="none": provides the theme without painting a background over the page.
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
