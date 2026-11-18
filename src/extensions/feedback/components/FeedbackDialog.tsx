import * as React from 'react';
import { DefaultButton, IButton, IButtonProps, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType, IDialogStyles } from '@fluentui/react/lib/Dialog';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { TextField } from '@fluentui/react/lib/TextField';
import { MailPreview } from './MailPreview';
import { SettingsView } from './SettingsView';
import { useClassNames } from './styles';
import { ISecurityPolicy } from '../config/securityConfig';
import {
  FeedbackErrors,
  FeedbackField,
  IFeedbackForm,
  IFeedbackOrigin,
  IPageInfo,
  IRecipientDefaults,
  ISiteSettings
} from '../models/IFeedback';
import { MailService } from '../services/MailService';
import { RecipientService } from '../services/RecipientService';
import { SettingsService } from '../services/SettingsService';
import { SetupService } from '../services/SetupService';
import { buildMail, IMail } from '../utils/mail';
import { FIELD_ORDER, hasErrors, validateForm } from '../utils/validation';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface IFeedbackDialogProps {
  origin: IFeedbackOrigin;
  policy: ISecurityPolicy;
  settingsService: SettingsService;
  mailService: MailService;
  recipientService: RecipientService;
  setupService: SetupService;
  /** Current user may open the settings (gear icon): at least "Design" permission on the site. */
  canConfigure: boolean;
  /** Current user may store entries in the settings list (once it is set up). */
  canSaveSettings: boolean;
  /** The page on which the dialog was opened (item of the Site Pages library). */
  page: IPageInfo;
  /** Pre-filled value of the URL field. */
  pageUrl: string;
  onDismiss: () => void;
}

/** Default width of the dialog; on wide screens it grows up to this width. */
const DIALOG_MAX_WIDTH = 1100;

/** Share of the window width the dialog takes (up to DIALOG_MAX_WIDTH). */
const DIALOG_WIDTH = '90vw';

/**
 * Responsive: 90 % of the window width up to 1100 px; below 480 px (phones) full screen.
 *
 * Passed as Dialog "styles" (not modalProps.styles): Fluent's Dialog sets "width: auto" in its own
 * "@media (min-width: 480px)" rule, which shrinks the dialog to its content. Dialog styles are merged
 * with these defaults, so the same media query here overrides it.
 */
const DIALOG_STYLES: Partial<IDialogStyles> = {
  main: {
    width: DIALOG_WIDTH,
    selectors: {
      '@media (min-width: 480px)': {
        width: DIALOG_WIDTH
      },
      '@media (max-width: 479px)': {
        width: '100vw',
        maxWidth: '100vw',
        minWidth: '100vw',
        height: '100%',
        maxHeight: '100%',
        borderRadius: 0
      }
    }
  }
};

/** Separator for several addresses in one field (multi-value lookup field). */
const EMAIL_SEPARATOR = '; ';

type Status = 'editing' | 'sending' | 'sent' | 'failed';

/** Anything with a focus() method: TextField and Dropdown expose it via componentRef. */
interface IFocusable { focus: () => void; }

/**
 * Dialog "Feedback geben / Inhalte melden": left the request, right the content responsibility
 * (recipients, pre-filled from the page's contact field and the web's property bag, editable).
 *
 * The component is mounted anew for every opening, so the form always starts fresh.
 */
export const FeedbackDialog: React.FunctionComponent<IFeedbackDialogProps> = (props: IFeedbackDialogProps) => {
  const {
    origin, policy, settingsService, mailService, recipientService, setupService,
    canConfigure, canSaveSettings, page, pageUrl, onDismiss
  } = props;
  const classNames = useClassNames();

  const [form, setForm] = React.useState<IFeedbackForm>(() => ({
    category: '',
    subject: '',
    url: pageUrl.substr(0, policy.maxTextLength),
    description: '',
    contactEmail: '',
    mailboxEmail: ''
  }));
  const [status, setStatus] = React.useState<Status>('editing');
  /** Errors are shown only after the first attempt to send, then they update while typing. */
  const [showErrors, setShowErrors] = React.useState(false);
  /** Settings of the site (code configuration and settings list); undefined while loading. */
  const [site, setSite] = React.useState<ISiteSettings | undefined>(undefined);
  /** undefined while the default recipients are loading. */
  const [defaults, setDefaults] = React.useState<IRecipientDefaults | undefined>(undefined);
  /** Incremented after a change in the settings view: reads settings and recipients again. */
  const [reloadCount, setReloadCount] = React.useState(0);
  /**
   * Form, preview of the mail (before sending) or settings view (gear icon).
   * The form keeps its input while the preview or the settings are shown.
   */
  const [view, setView] = React.useState<'form' | 'preview' | 'settings'>('form');
  /** The mail shown in the preview; exactly this mail is sent. */
  const [preview, setPreview] = React.useState<IMail | undefined>(undefined);

  const fieldRefs = React.useRef<{ [K in FeedbackField]?: IFocusable | null }>({});
  const closeButton = React.useRef<IButton>(null);
  const sendButton = React.useRef<IButton>(null);
  const mounted = React.useRef(true);
  React.useEffect(() => () => { mounted.current = false; }, []);
  /** Addresses last pre-filled: after a change in the settings an unchanged field takes the new value. */
  const prefilled = React.useRef({ contactEmail: '', mailboxEmail: '' });

  // Load the settings of the site, then the contact person of the page and the shared mailbox of the web.
  // A field the user has typed into is not overwritten.
  React.useEffect(() => {
    settingsService.load()
      .then((loadedSite: ISiteSettings) => {
        if (!mounted.current) {
          return undefined;
        }
        setSite(loadedSite);
        return recipientService.load(page, loadedSite);
      })
      .then((loaded: IRecipientDefaults | undefined) => {
        if (!mounted.current || !loaded) {
          return;
        }
        setDefaults(loaded);
        const next = { contactEmail: loaded.contact.emails.join(EMAIL_SEPARATOR), mailboxEmail: loaded.mailbox.emails.join(EMAIL_SEPARATOR) };
        const previous = prefilled.current;
        prefilled.current = next;
        setForm((current: IFeedbackForm) => ({
          ...current,
          contactEmail: !current.contactEmail || current.contactEmail === previous.contactEmail ? next.contactEmail : current.contactEmail,
          mailboxEmail: !current.mailboxEmail || current.mailboxEmail === previous.mailboxEmail ? next.mailboxEmail : current.mailboxEmail
        }));
      });
  }, [settingsService, recipientService, page, reloadCount]);

  const onSettingsChanged = (): void => {
    settingsService.reset();
    recipientService.reset();
    setReloadCount((count: number) => count + 1);
  };

  /** Until the settings list is read, the code configuration applies (categories are not selectable yet). */
  const settings = site ? site.settings : settingsService.base;
  const loading = !defaults;
  /** Hint below an address field: loading, or nothing stored in SharePoint. */
  const hintFor = (found: boolean, missingText: string): string | undefined =>
    loading ? strings.loadingRecipients : (found ? undefined : missingText);

  const errors: FeedbackErrors = showErrors ? validateForm(form, settings.categories, policy) : {};
  const busy = status === 'sending';
  const done = status === 'sent';

  // After sending successfully only the confirmation and "Schließen" remain: move the focus there.
  React.useEffect(() => {
    if (done && closeButton.current) {
      closeButton.current.focus();
    }
  }, [done]);

  // The preview replaces the form: the focus goes to "Senden".
  React.useEffect(() => {
    if (view === 'preview' && sendButton.current) {
      sendButton.current.focus();
    }
  }, [view]);

  const setField = (field: FeedbackField) => (_event: unknown, value?: string): void => {
    setForm((current: IFeedbackForm) => ({ ...current, [field]: value || '' }));
    if (status === 'failed') {
      setStatus('editing');
    }
  };

  const onCategoryChange = (_event: unknown, option?: IDropdownOption): void => {
    setField('category')(undefined, option ? String(option.key) : '');
  };

  const refFor = (field: FeedbackField) => (component: IFocusable | null): void => {
    fieldRefs.current[field] = component;
  };

  /** Checks the form and shows the mail as it will be sent. */
  const showPreview = (): void => {
    if (busy || done || !site) {
      return;
    }
    const currentErrors = validateForm(form, settings.categories, policy);
    if (hasErrors(currentErrors)) {
      setShowErrors(true);
      const first = FIELD_ORDER.filter((field: FeedbackField) => !!currentErrors[field])[0];
      const target = fieldRefs.current[first];
      if (target) {
        target.focus();
      }
      return;
    }
    setPreview(buildMail(form, origin, settings.mail, new Date()));
    setStatus('editing');
    setView('preview');
  };

  /** Sends the mail of the preview. */
  const send = (): void => {
    if (busy || done || !preview) {
      return;
    }
    setStatus('sending');
    mailService.send(preview).then(
      () => { if (mounted.current) { setStatus('sent'); } },
      () => { if (mounted.current) { setStatus('failed'); } }
    );
  };

  const onFormSubmit = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    showPreview();
  };

  const options: IDropdownOption[] = settings.categories.map((category: string) => ({ key: category, text: category }));

  const showSettings = canConfigure && view === 'settings';
  const showPreviewView = view === 'preview' && !!preview;
  // Gear icon next to the close button, only for users who may configure. The settings view has its own
  // "Back" button (it asks before discarding unsaved changes), so the gear is shown in the form only.
  const topButtons: IButtonProps[] = canConfigure && view === 'form' && !busy && !done ? [{
    iconProps: { iconName: 'Settings' },
    title: strings.settingsOpen,
    ariaLabel: strings.settingsOpen,
    onClick: () => setView('settings')
  }] : [];

  return (
    <Dialog
      hidden={false}
      onDismiss={busy ? undefined : onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: showSettings ? strings.settingsTitle : strings.dialogTitle,
        showCloseButton: !busy,
        closeButtonAriaLabel: strings.close,
        topButtonsProps: topButtons
      }}
      // isBlocking: a click next to the dialog must not discard the input.
      modalProps={{ isBlocking: true }}
      styles={DIALOG_STYLES}
      minWidth={300}
      maxWidth={DIALOG_MAX_WIDTH}
    >
      {done ? (
        <>
          <MessageBar messageBarType={MessageBarType.success} className={classNames.message}>{strings.sendSucceeded}</MessageBar>
          <DialogFooter>
            <DefaultButton componentRef={closeButton} text={strings.close} onClick={onDismiss} />
          </DialogFooter>
        </>
      ) : showPreviewView && preview ? (
        <>
          {status === 'failed' && (
            <MessageBar messageBarType={MessageBarType.error} className={classNames.message}>{strings.sendFailed}</MessageBar>
          )}
          <MailPreview mail={preview} />
          <DialogFooter>
            <PrimaryButton componentRef={sendButton} onClick={send} disabled={busy} text={busy ? strings.sending : strings.previewSend}>
              {busy && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
            </PrimaryButton>
            <DefaultButton text={strings.previewBack} onClick={() => { setStatus('editing'); setView('form'); }} disabled={busy} />
            <DefaultButton text={strings.cancel} onClick={onDismiss} disabled={busy} />
          </DialogFooter>
        </>
      ) : showSettings ? (
        // Not inside the feedback form: Enter in a settings field must not send the report.
        <SettingsView
          site={site}
          defaults={defaults}
          base={settingsService.base}
          page={page}
          canSaveSettings={canSaveSettings}
          policy={policy}
          setupService={setupService}
          webUrl={origin.webUrl}
          onChanged={onSettingsChanged}
          onBack={() => setView('form')}
        />
      ) : (
        <>
          {/* noValidate: the browser's own validation bubbles would compete with the Fluent messages.
              The footer is part of the form: Enter in a single-line field sends, i.e. shows the preview first (also in IE11). */}
          <form noValidate onSubmit={onFormSubmit}>
          {showErrors && hasErrors(errors) && (
            <MessageBar messageBarType={MessageBarType.warning} className={classNames.message}>{strings.errorForm}</MessageBar>
          )}
          <div className={classNames.columns}>
            <fieldset className={classNames.column} disabled={busy}>
              <legend className={classNames.sectionTitle}>{strings.sectionRequest}</legend>
              <Dropdown
                componentRef={refFor('category')}
                className={classNames.field}
                label={strings.categoryLabel}
                placeholder={site ? strings.categoryPlaceholder : strings.loadingRecipients}
                options={options}
                selectedKey={form.category || null}
                onChange={onCategoryChange}
                errorMessage={errors.category}
                required
                disabled={busy || !site}
              />
              <TextField
                componentRef={refFor('subject')}
                className={classNames.field}
                label={strings.subjectLabel}
                placeholder={strings.subjectPlaceholder}
                value={form.subject}
                onChange={setField('subject')}
                errorMessage={errors.subject}
                maxLength={policy.maxTextLength}
                required
              />
              <TextField
                componentRef={refFor('url')}
                className={classNames.field}
                label={strings.urlLabel}
                placeholder={strings.urlPlaceholder}
                value={form.url}
                onChange={setField('url')}
                errorMessage={errors.url}
                maxLength={policy.maxTextLength}
                type="url"
                inputMode="url"
                required
              />
              <TextField
                componentRef={refFor('description')}
                className={classNames.field}
                label={strings.descriptionLabel}
                placeholder={strings.descriptionPlaceholder}
                value={form.description}
                onChange={setField('description')}
                errorMessage={errors.description}
                maxLength={policy.maxDescriptionLength}
                multiline
                rows={8}
                resizable={false}
                required
              />
            </fieldset>
            <fieldset className={classNames.column} disabled={busy}>
              <legend className={classNames.sectionTitle}>{strings.sectionResponsibility}</legend>
              <TextField
                componentRef={refFor('contactEmail')}
                className={classNames.field}
                label={strings.contactLabel}
                placeholder={strings.contactPlaceholder}
                value={form.contactEmail}
                onChange={setField('contactEmail')}
                errorMessage={errors.contactEmail}
                description={hintFor(!!defaults && defaults.contact.status === 'ok', strings.contactMissing)}
                maxLength={policy.maxTextLength}
                inputMode="email"
                autoComplete="off"
                required
              />
              <TextField
                componentRef={refFor('mailboxEmail')}
                className={classNames.field}
                label={strings.mailboxLabel}
                placeholder={strings.mailboxPlaceholder}
                value={form.mailboxEmail}
                onChange={setField('mailboxEmail')}
                errorMessage={errors.mailboxEmail}
                description={hintFor(!!defaults && defaults.mailbox.status === 'ok', strings.mailboxMissing)}
                maxLength={policy.maxTextLength}
                inputMode="email"
                autoComplete="off"
              />
              <p className={classNames.hint}>{strings.responsibilityHint}</p>
            </fieldset>
          </div>
          <DialogFooter>
            <PrimaryButton type="submit" text={strings.send} />
            <DefaultButton text={strings.cancel} onClick={onDismiss} disabled={busy} />
          </DialogFooter>
          </form>
        </>
      )}
    </Dialog>
  );
};
