import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { DialogFooter } from '@fluentui/react/lib/Dialog';
import { Link } from '@fluentui/react/lib/Link';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ITextFieldProps, TextField } from '@fluentui/react/lib/TextField';
import { useClassNames, useSetupClassNames } from './styles';
import { ISecurityPolicy } from '../config/securityConfig';
import { formatText } from '../loc/template';
import {
  IFeedbackSettings,
  IPageInfo,
  IRecipientDefaults,
  ISettingKeys,
  ISiteSettings,
  IStoredValues,
  RecipientStatus
} from '../models/IFeedback';
import { SetupError, SetupService } from '../services/SetupService';
import { MAIL_PLACEHOLDERS } from '../utils/mail';
import { singleLine, splitLines } from '../utils/sanitize';
import { checkCategories, checkColumnName, checkListName, checkMailTemplate } from '../utils/setupValidation';
import { checkEmailList, splitEmails } from '../utils/validation';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface ISetupViewProps {
  /** undefined while loading */
  defaults: IRecipientDefaults | undefined;
  /** Settings of the site (code configuration and settings list); undefined while loading */
  site: ISiteSettings | undefined;
  /** Code configuration: defaults shown as placeholders. */
  base: IFeedbackSettings;
  page: IPageInfo;
  /** Current user may store entries in the settings list (design permission). */
  canSaveSettings: boolean;
  /** Current user may store the shared mailbox (settings list: design permission, property bag: site owners). */
  canSaveMailbox: boolean;
  policy: ISecurityPolicy;
  setupService: SetupService;
  webUrl: string;
  /** Called after a successful change: the dialog reads settings and recipients again. */
  onChanged: () => void;
  /** Back to the feedback form. */
  onBack: () => void;
}

/** Fields stored in the settings list (one entry each, see config settingsList.keys). */
type SettingField = keyof ISettingKeys;

/** All editable fields: the settings and the shared mailbox (settings list or property bag). */
type Field = SettingField | 'mailbox';
type Values = { [F in Field]: string };
type Flags = { [F in Field]?: boolean };
type Errors = { [F in Field]?: string };

const EMPTY: Values = { contactList: '', contactColumn: '', categories: '', mailSubject: '', mailBody: '', mailbox: '' };

/** Contact states that "Set up" can fix (create list, e-mail column, lookup column). */
const CONTACT_FIXABLE: RecipientStatus[] = ['missing', 'lookupEmailMissing'];

/** True if something is missing or wrong: used for the hint in the form. */
export function isSetupIncomplete(defaults: IRecipientDefaults | undefined): boolean {
  return !!defaults && (
    CONTACT_FIXABLE.indexOf(defaults.contact.status) >= 0 || defaults.contact.status === 'wrongType'
    || defaults.mailbox.status === 'missing' || defaults.mailbox.status === 'invalid'
  );
}

type Tone = 'ok' | 'problem' | 'neutral';

/** Message for a failed setup action: missing permission or a general failure. */
function failureText(error: unknown): string {
  return error instanceof SetupError && error.accessDenied ? strings.setupAccessDenied : strings.setupFailed;
}

/** Values of the fields as stored: entries of the settings list, the shared mailbox as read for the form. */
function storedValues(site: ISiteSettings | undefined, defaults: IRecipientDefaults | undefined): Values {
  const stored: IStoredValues = site && site.stored ? site.stored : {};
  const keys = site ? site.settings.settingsList.keys : undefined;
  const entry = (field: SettingField): string => (keys && typeof stored[keys[field]] === 'string' ? stored[keys[field]] : '');
  return {
    contactList: entry('contactList'),
    contactColumn: entry('contactColumn'),
    categories: entry('categories'),
    mailSubject: entry('mailSubject'),
    mailBody: entry('mailBody'),
    mailbox: defaults ? defaults.mailbox.emails.join('; ') : ''
  };
}

/** Value as it is stored: trimmed, categories one per line without empty lines. */
function normalize(field: Field, value: string): string {
  switch (field) {
    case 'categories': return splitLines(value).map(singleLine).filter((line: string) => !!line).join('\n');
    case 'mailSubject': return singleLine(value);
    case 'mailbox': return splitEmails(value).join('; ');
    default: return value.trim();
  }
}

/**
 * Setup view of the dialog (gear icon, only for users with at least "Design" permission on the site).
 * A plain form in two columns like the feedback form:
 *  - contact person: list and lookup column (settings list), status of column, list and the page's contact;
 *    "Set up" creates what is missing
 *  - shared mailbox (settings list or property bag)
 *  - form: categories (settings list)
 *  - notification: mail subject and text (settings list)
 * "Save" stores all changed fields. An empty field = the default from the code (shown as placeholder).
 */
export const SetupView: React.FunctionComponent<ISetupViewProps> = (props: ISetupViewProps) => {
  const { defaults, site, base, page, canSaveSettings, canSaveMailbox, policy, setupService, webUrl, onChanged, onBack } = props;
  const formClassNames = useClassNames();
  const classNames = useSetupClassNames();

  const [contactBusy, setContactBusy] = React.useState(false);
  const [contactMessage, setContactMessage] = React.useState<{ text: string; tone: Tone } | undefined>(undefined);

  const [values, setValues] = React.useState<Values>(EMPTY);
  const [dirty, setDirty] = React.useState<Flags>({});
  const [errors, setErrors] = React.useState<Errors>({});
  const [saving, setSaving] = React.useState(false);
  const [saveMessage, setSaveMessage] = React.useState<{ text: string; tone: Tone } | undefined>(undefined);
  /** Saved, waiting for the dialog to read the new values: the fields keep their input until then. */
  const [reloading, setReloading] = React.useState(false);

  const mounted = React.useRef(true);
  React.useEffect(() => () => { mounted.current = false; }, []);

  // Pre-fill the fields with the stored values (again after saving), except fields the user is editing.
  const stored = storedValues(site, defaults);
  const storedKey = JSON.stringify(stored);
  React.useEffect(() => {
    setValues((current: Values) => {
      const next = { ...current };
      (Object.keys(EMPTY) as Field[]).forEach((field: Field) => {
        if (!dirty[field]) {
          next[field] = stored[field];
        }
      });
      return next;
    });
    // storedKey stands for stored (a new object on every render).
  }, [storedKey, dirty]);

  // The dialog reads settings and recipients after saving; the recipients come last.
  React.useEffect(() => {
    if (reloading) {
      setReloading(false);
      setDirty({});
    }
  }, [defaults]);

  const settings = site ? site.settings : base;
  const contact = defaults ? defaults.contact : undefined;
  const mailbox = defaults ? defaults.mailbox : undefined;
  const baseUrl = webUrl.replace(/\/+$/, '');
  const emailColumn = settings.contactEmailColumn.displayName;
  const editable = (field: Field): boolean => (field === 'mailbox' ? canSaveMailbox : canSaveSettings);
  const anyDirty = (Object.keys(dirty) as Field[]).some((field: Field) => !!dirty[field]);

  const setField = (field: Field, value: string): void => {
    setValues((current: Values) => ({ ...current, [field]: value }));
    setDirty((current: Flags) => ({ ...current, [field]: true }));
    setErrors((current: Errors) => ({ ...current, [field]: undefined }));
    setSaveMessage(undefined);
  };

  const validate = (field: Field, value: string): string => {
    switch (field) {
      case 'contactList': return checkListName(value);
      case 'contactColumn': return checkColumnName(value);
      case 'categories': return checkCategories(value, policy);
      case 'mailSubject': return checkMailTemplate(value, policy.maxSubjectLength);
      case 'mailBody': return checkMailTemplate(value, policy.maxMailTemplateLength);
      default: return value.trim() ? checkEmailList(value, policy) : strings.errorRequired;
    }
  };

  const save = (): void => {
    if (saving || !anyDirty) {
      return;
    }
    const changed = (Object.keys(EMPTY) as Field[]).filter((field: Field) => dirty[field] && editable(field));
    const nextErrors: Errors = {};
    changed.forEach((field: Field) => {
      const error = validate(field, values[field]);
      if (error) { nextErrors[field] = error; }
    });
    if (changed.some((field: Field) => !!nextErrors[field])) {
      setErrors(nextErrors);
      return;
    }

    const entries: IStoredValues = {};
    let propertyBagMailbox: string | undefined;
    changed.forEach((field: Field) => {
      const value = normalize(field, values[field]);
      if (field !== 'mailbox') {
        entries[settings.settingsList.keys[field]] = value;
      } else if (settings.mailbox.storage === 'list') {
        entries[settings.mailbox.key] = value;
      } else {
        propertyBagMailbox = value;
      }
    });

    setSaving(true);
    setSaveMessage(undefined);
    setupService.saveValues(entries)
      .then(() => (propertyBagMailbox !== undefined ? setupService.setMailbox(propertyBagMailbox) : undefined))
      .then(
        () => {
          if (!mounted.current) { return; }
          setSaving(false);
          setSaveMessage({ text: strings.setupSaved, tone: 'ok' });
          setReloading(true);
          onChanged();
        },
        (error: unknown) => {
          if (!mounted.current) { return; }
          setSaving(false);
          setSaveMessage({ text: failureText(error), tone: 'problem' });
        }
      );
  };

  const setupContact = (): void => {
    if (!contact) { return; }
    setContactBusy(true);
    setContactMessage(undefined);
    setupService.setupContact(settings, page.listId, contact.status === 'lookupEmailMissing' ? contact.lookupListId : undefined).then(
      () => {
        if (!mounted.current) { return; }
        setContactBusy(false);
        setContactMessage({ text: strings.setupContactDone, tone: 'ok' });
        onChanged();
      },
      (error: unknown) => {
        if (!mounted.current) { return; }
        setContactBusy(false);
        setContactMessage({ text: failureText(error), tone: 'problem' });
      }
    );
  };

  /** Copies the default template of the current language into the fields, as a starting point. */
  const insertDefaultTemplate = (): void => {
    setField('mailSubject', base.mail.subject);
    setField('mailBody', base.mail.body);
  };

  /** Props shared by all text fields of the settings. Enter in a single-line field saves. */
  const fieldProps = (field: Field, multiline?: boolean): Partial<ITextFieldProps> => ({
    className: formClassNames.field,
    value: values[field],
    onChange: (_e: unknown, value?: string) => setField(field, value || ''),
    onKeyDown: multiline ? undefined : (event: React.KeyboardEvent) => {
      if (event.key === 'Enter' && editable(field)) {
        event.preventDefault();
        save();
      }
    },
    errorMessage: errors[field],
    readOnly: !editable(field),
    disabled: saving,
    autoComplete: 'off'
  });

  /** One line "label: value" with an optional status tone. */
  const row = (label: string, value: React.ReactNode, tone: Tone = 'neutral'): React.ReactNode => (
    <div className={classNames.row}>
      <div className={classNames.label}>{label}</div>
      <div className={tone === 'ok' ? classNames.ok : tone === 'problem' ? classNames.problem : classNames.value}>{value}</div>
    </div>
  );

  const columnStatus = (): [string, Tone] => {
    switch (contact ? contact.status : 'error') {
      case 'missing': return [strings.setupStatusMissing, 'problem'];
      case 'wrongType': return [strings.setupStatusWrongType, 'problem'];
      case 'lookupEmailMissing': return [formatText(strings.setupStatusEmailMissing, emailColumn), 'problem'];
      case 'notConfigured': return [strings.setupNotConfigured, 'problem'];
      case 'error': return [strings.setupStatusUnknown, 'problem'];
      default: return [strings.setupStatusPresent, 'ok'];
    }
  };

  const newWindow = { target: '_blank', rel: 'noopener noreferrer' };
  const editFormUrl = `${baseUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${encodeURIComponent('{' + page.listId + '}')}&ID=${page.itemId}`;
  const columnExists = !!contact && ['ok', 'empty', 'invalid', 'lookupEmailMissing'].indexOf(contact.status) >= 0;

  const pageContact = (): [React.ReactNode, Tone] => {
    if (!contact || !columnExists || contact.status === 'lookupEmailMissing') {
      return ['–', 'neutral'];
    }
    const text = contact.status === 'ok' ? contact.emails.join('; ')
      : contact.status === 'invalid' ? strings.setupContactInvalid : strings.setupPageContactNone;
    return [
      <>
        {text}
        {page.canEditItems && <>{' · '}<Link href={editFormUrl} {...newWindow}>{strings.setupEditPageProperties}</Link></>}
      </>,
      contact.status === 'ok' ? 'neutral' : 'problem'
    ];
  };

  const mailboxStatus = (): string | undefined => {
    if (!mailbox) { return undefined; }
    if (mailbox.status === 'invalid') { return strings.setupMailboxInvalid; }
    if (mailbox.status === 'error') { return strings.setupReadFailed; }
    return undefined;
  };

  const [columnText, columnTone] = columnStatus();
  const [pageContactNode, pageContactTone] = pageContact();
  const message = (text: string, tone: Tone): React.ReactNode => (
    <p className={tone === 'ok' ? classNames.ok : classNames.problem} role={tone === 'ok' ? 'status' : 'alert'}>{text}</p>
  );
  const placeholders = MAIL_PLACEHOLDERS.map((name: string) => `{{${name}}}`).join(' ');

  return (
    <>
      {!defaults || !site ? (
        <Spinner label={strings.loadingRecipients} className={classNames.loading} />
      ) : (
        <>
          {!site.stored && message(strings.setupSettingsReadFailed, 'problem')}
          <div className={formClassNames.columns}>
            <section className={formClassNames.column} aria-labelledby="spfx-feedback-setup-contact">
              <h3 id="spfx-feedback-setup-contact" className={formClassNames.sectionTitle}>{strings.setupContactHeading}</h3>
              <TextField
                {...fieldProps('contactList')}
                label={strings.setupContactListField}
                placeholder={base.contactList.url.substring('Lists/'.length)}
                description={strings.setupContactListDescription}
              />
              <TextField
                {...fieldProps('contactColumn')}
                label={strings.setupContactColumnField}
                placeholder={base.contactColumn.internalName}
                description={strings.setupContactColumnDescription}
              />
              {row(strings.setupColumnLabel, <>{settings.contactColumn.displayName} – {columnText}</>, columnTone)}
              {row(strings.setupListLabel,
                <Link href={`${baseUrl}/${settings.contactList.url}`} {...newWindow}>{settings.contactList.title}</Link>)}
              {row(strings.setupPageContactLabel, pageContactNode, pageContactTone)}
              {contact && CONTACT_FIXABLE.indexOf(contact.status) >= 0 && (
                <div className={classNames.action}>
                  <p className={classNames.hint}>
                    {formatText(strings.setupContactMissing, settings.contactList.title, emailColumn, settings.contactColumn.displayName)}
                  </p>
                  {/* Set up works with the saved list and column: unsaved changes must be saved first. */}
                  <DefaultButton
                    text={strings.setupContactSetup}
                    onClick={setupContact}
                    disabled={contactBusy || saving || !!dirty.contactList || !!dirty.contactColumn}
                  >
                    {contactBusy && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
                  </DefaultButton>
                </div>
              )}
              {contactMessage && message(contactMessage.text, contactMessage.tone)}
            </section>

            <section className={formClassNames.column} aria-labelledby="spfx-feedback-setup-mailbox">
              <h3 id="spfx-feedback-setup-mailbox" className={formClassNames.sectionTitle}>{strings.setupMailboxHeading}</h3>
              <TextField
                {...fieldProps('mailbox')}
                label={strings.setupMailboxLabel}
                placeholder={strings.mailboxPlaceholder}
                description={canSaveMailbox
                  ? (settings.mailbox.storage === 'list'
                    ? formatText(strings.setupMailboxDescriptionList, settings.settingsList.title)
                    : formatText(strings.setupMailboxDescription, settings.mailbox.key))
                  : strings.setupMailboxNeedsOwner}
                maxLength={policy.maxRecipientsPerField * (policy.maxEmailLength + 2)}
                inputMode="email"
              />
              {mailboxStatus() && message(mailboxStatus() as string, 'problem')}
            </section>
          </div>

          <div className={formClassNames.columns}>
            <section className={formClassNames.column} aria-labelledby="spfx-feedback-setup-form">
              <h3 id="spfx-feedback-setup-form" className={formClassNames.sectionTitle}>{strings.setupFormHeading}</h3>
              <TextField
                {...fieldProps('categories', true)}
                label={strings.setupCategoriesLabel}
                placeholder={base.categories.join('\n')}
                description={strings.setupCategoriesDescription}
                multiline
                rows={8}
              />
            </section>

            <section className={formClassNames.column} aria-labelledby="spfx-feedback-setup-mail">
              <h3 id="spfx-feedback-setup-mail" className={formClassNames.sectionTitle}>{strings.setupMailHeading}</h3>
              <TextField
                {...fieldProps('mailSubject')}
                label={strings.setupMailSubjectLabel}
                placeholder={base.mail.subject}
                maxLength={policy.maxSubjectLength}
              />
              <TextField
                {...fieldProps('mailBody', true)}
                label={strings.setupMailBodyLabel}
                placeholder={strings.setupMailBodyPlaceholder}
                description={formatText(strings.setupMailDescription, placeholders)}
                maxLength={policy.maxMailTemplateLength}
                multiline
                rows={8}
              />
              {canSaveSettings && (
                <p className={classNames.hint}>
                  <Link onClick={insertDefaultTemplate} disabled={saving}>{strings.setupInsertDefault}</Link>
                </p>
              )}
            </section>
          </div>
          {saveMessage && message(saveMessage.text, saveMessage.tone)}
        </>
      )}
      <p className={classNames.visibility}>
        {strings.setupVisibility} {formatText(strings.setupSettingsStored, base.settingsList.title)}
      </p>
      <DialogFooter>
        {(canSaveSettings || canSaveMailbox) && (
          <PrimaryButton text={strings.setupSave} onClick={save} disabled={!defaults || !site || saving || !anyDirty}>
            {saving && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
          </PrimaryButton>
        )}
        <DefaultButton text={strings.setupBack} onClick={onBack} disabled={saving || contactBusy} />
      </DialogFooter>
    </>
  );
};
