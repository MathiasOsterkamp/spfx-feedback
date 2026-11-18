import * as React from 'react';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { setIconOptions } from '@fluentui/react/lib/Styling';
import * as ReactDom from 'react-dom';
import { Feedback, IFeedbackProps } from './components/Feedback';
import { FEEDBACK_CONFIG } from './config/feedbackConfig';
import { createSecurityPolicy } from './config/securityConfig';
import { IPageInfo } from './models/IFeedback';
import { MailService } from './services/MailService';
import { RecipientService } from './services/RecipientService';
import { SettingsService } from './services/SettingsService';
import { SetupService } from './services/SetupService';
import { configureLogging, logInfo, logVerbose, logWarning, LogLevelName, sanitizeLogLevel } from './utils/logging';
import { createSettings } from './utils/sanitize';
import { createSharePointTheme } from './utils/theme';
import { asRecord } from './utils/untyped';

/** ClientSideComponentProperties (sharepoint/assets/elements.xml); all else: config/feedbackConfig.ts and settings list. */
export interface IFeedbackApplicationCustomizerProperties {
  /** Minimum level written to the browser console. */
  logLevel?: LogLevelName;
}

/** Base template of the "Site Pages" library. */
const SITE_PAGES_TEMPLATE = 119;

// eslint-disable-next-line @typescript-eslint/no-require-imports -- an import would pull config/ into the compiler's rootDir and shift the output in lib/
const SOLUTION = asRecord(asRecord(require('../../../config/package-solution.json')).solution);

/** Prevents a second button if the custom action is registered twice. */
const CONTAINER_ID = 'spfx-feedback-button';

export default class FeedbackApplicationCustomizer extends BaseApplicationCustomizer<IFeedbackApplicationCustomizerProperties> {
  private container: HTMLElement | undefined;
  /** Everything except the current page, which changes on navigation. */
  private baseProps: Pick<IFeedbackProps, Exclude<keyof IFeedbackProps, 'page'>> | undefined;

  protected onInit(): Promise<void> {
    configureLogging(sanitizeLogLevel(asRecord(this.properties).logLevel));
    logInfo('Extension initialized', { solution: SOLUTION.name, version: SOLUTION.version });

    if (document.getElementById(CONTAINER_ID)) {
      logWarning('Feedback button already present, the extension is probably registered twice');
      return Promise.resolve();
    }

    // SharePoint sometimes already registers icons with the same name -> suppress warnings.
    setIconOptions({ disableWarnings: true });
    // On-premises without internet: host the font files yourself and call initializeIcons('https://.../fonts/').
    initializeIcons();

    const policy = createSecurityPolicy();
    // The settings list is read only when the dialog opens (SettingsService).
    const settings = createSettings(FEEDBACK_CONFIG, this.context.pageContext.cultureInfo.currentUICultureName, policy);
    const web = this.context.pageContext.web;
    const user = this.context.pageContext.user;
    logVerbose('Settings (code)', settings);

    const http = this.context.spHttpClient;
    // Settings list: create it and edit its entries -> "Design" permission (ManageLists, EditListItems).
    const canSaveSettings = web.permissions.hasPermission(SPPermission.manageLists) && web.permissions.hasPermission(SPPermission.editListItems);
    this.baseProps = {
      theme: createSharePointTheme(),
      origin: {
        userName: user.displayName || user.loginName || '',
        userEmail: user.email || '',
        webTitle: web.title || '',
        webUrl: web.absoluteUrl
      },
      policy,
      settingsService: new SettingsService(http, web.absoluteUrl, web.serverRelativeUrl, settings, policy),
      mailService: new MailService(http, web.absoluteUrl),
      recipientService: new RecipientService(http, web.absoluteUrl, policy),
      setupService: new SetupService(http, web.absoluteUrl, web.serverRelativeUrl, settings),
      // Gear icon: "Design" permission level (includes ManageLists).
      canConfigure: web.permissions.hasPermission(SPPermission.manageLists),
      canSaveSettings
    };

    // Own container at the end of the body: the floating button is independent of the SharePoint placeholders.
    this.container = document.createElement('div');
    this.container.id = CONTAINER_ID;
    document.body.appendChild(this.container);

    // Modern pages navigate without reloading (e.g. from a page to a list): check again every time.
    this.context.application.navigatedEvent.add(this, this.render);
    this.render();

    return Promise.resolve();
  }

  protected onDispose(): void {
    this.context.application.navigatedEvent.remove(this, this.render);
    if (this.container) {
      ReactDom.unmountComponentAtNode(this.container);
      if (this.container.parentNode) {
        this.container.parentNode.removeChild(this.container);
      }
      this.container = undefined;
    }
  }

  private render(): void {
    if (!this.container || !this.baseProps) {
      return;
    }
    const page = this.currentSitePage();
    logVerbose('Render', { page });
    ReactDom.render(React.createElement(Feedback, { ...this.baseProps, page }), this.container);
  }

  /**
   * Only items of Site Pages. SPFx 1.4.1 does not expose the list base template: it comes from the
   * legacy page context, otherwise the library URL ".../SitePages" decides.
   */
  private currentSitePage(): IPageInfo | undefined {
    const pageContext = this.context.pageContext;
    const legacy = asRecord(pageContext.legacyPageContext);
    const list = pageContext.list;
    const itemId = pageContext.listItem ? pageContext.listItem.id : legacy.pageItemId;
    if (!list || typeof itemId !== 'number' || itemId <= 0) {
      return undefined;
    }
    const template = legacy.listBaseTemplate;
    const isSitePages = typeof template === 'number'
      ? template === SITE_PAGES_TEMPLATE
      : /\/SitePages$/i.test(list.serverRelativeUrl);
    if (!isSitePages) {
      return undefined;
    }
    // Only decides which setup actions are offered (SharePoint checks on the server):
    // the contact list needs ManageLists on the web, the lookup column ManageLists on Site Pages.
    const web = pageContext.web;
    return {
      listId: list.id.toString(),
      itemId,
      canManageList: web.permissions.hasPermission(SPPermission.manageLists) && list.permissions.hasPermission(SPPermission.manageLists),
      canEditItems: list.permissions.hasPermission(SPPermission.editListItems)
    };
  }
}
