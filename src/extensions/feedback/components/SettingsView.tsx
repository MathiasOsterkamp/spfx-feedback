import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { DialogFooter } from '@fluentui/react/lib/Dialog';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ITextFieldProps, TextField } from '@fluentui/react/lib/TextField';
import { useClassNames, useSettingsClassNames } from './styles';
import { ISecurityPolicy } from '../config/securityConfig';
import { formatText } from '../loc/template';
import {
  IFeedbackSettings,
  ILanguageSettings,
  IListRef,
  IPageInfo,
  IRecipientDefaults,
  ISettingKeys,
  ISetupStatus,
  ISiteSettings,
  IStoredValues,
  SetupPartStatus
} from '../models/IFeedback';
import { canSetUp, IContactListOption, ILookupColumnOption, isSetupComplete, SetupError, SetupService } from '../services/SetupService';
import { MAIL_PLACEHOLDERS } from '../utils/mail';
import { languageKey, singleLine, splitLines } from '../utils/sanitize';
import { checkCategories, checkColumnName, checkListName, checkMailTemplate } from '../utils/setupValidation';
import { checkAddressField, splitEmails } from '../utils/validation';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface ISettingsViewProps {
  site: ISiteSettings | undefined;
  defaults: IRecipientDefaults | undefined;
  base: IFeedbackSettings;
  page: IPageInfo;
  canSaveSettings: boolean;
  policy: ISecurityPolicy;
  setupService: SetupService;
  webUrl: string;
  onChanged: () => void;
  onBack: () => void;
}

type Tab = 'recipients' | 'texts' | 'setup';

type LanguageField = 'categories' | 'mailSubject' | 'mailBody';
type SiteField = Exclude<keyof ISettingKeys, LanguageField> | 'mailbox';
type Field = SiteField | LanguageField;

const SITE_FIELDS: SiteField[] = ['contactList', 'contactColumn', 'mailbox'];
const LANGUAGE_FIELDS: LanguageField[] = ['categories', 'mailSubject', 'mailBody'];

type Values = { [id: string]: string };
type Flags = { [id: string]: boolean | undefined };
type Errors = { [id: string]: string | undefined };

type Tone = 'ok' | 'problem';
interface IMessage { text: string; tone: Tone }

function fieldId(field: Field, language?: string): string {
  return language ? `${field}|${language}` : field;
}

function parseId(id: string): { field: Field; language: string | undefined } {
  const [field, language] = id.split('|');
  return { field: field as Field, language };
}

function allIds(languages: ILanguageSettings[]): string[] {
  return languages.reduce((ids: string[], language: ILanguageSettings) =>
    ids.concat(LANGUAGE_FIELDS.map((field: LanguageField) => fieldId(field, language.code))), SITE_FIELDS.slice());
}

function tabOf(id: string): Tab {
  return parseId(id).language ? 'texts' : 'recipients';
}

function entryKey(keys: ISettingKeys, id: string): string {
  const { field, language } = parseId(id);
  return field === 'mailbox' ? '' : (language ? languageKey(keys[field], language) : keys[field]);
}

function normalize(field: Field, value: string): string {
  switch (field) {
    case 'categories': return splitLines(value).map(singleLine).filter((line: string) => !!line).join('\n');
    case 'mailSubject': return singleLine(value);
    case 'mailbox': return splitEmails(value).join('; ');
    default: return value.trim();
  }
}

function failureText(error: unknown, general: string): string {
  return error instanceof SetupError && error.accessDenied ? strings.settingsAccessDenied : general;
}

const NEW_WINDOW = { target: '_blank', rel: 'noopener noreferrer' };

export const SettingsView: React.FunctionComponent<ISettingsViewProps> = (props: ISettingsViewProps) => {
  const { site, defaults, base, page, canSaveSettings, policy, setupService, webUrl, onChanged, onBack } = props;
  const formClassNames = useClassNames();
  const classNames = useSettingsClassNames();

  const settings = site ? site.settings : base;
  const languages = base.languages;
  const ids = allIds(languages);
  const contactListSelectable = !!base.settingsList.keys.contactList;
  const contactColumnSelectable = !!base.settingsList.keys.contactColumn;

  const [tab, setTab] = React.useState<Tab>('recipients');
  const tabChosen = React.useRef(false);
  const [languageCode, setLanguageCode] = React.useState(base.language);
  const language = languages.filter((entry: ILanguageSettings) => entry.code === languageCode)[0] || languages[0];

  const [status, setStatus] = React.useState<ISetupStatus | null | undefined>(undefined);
  const [checking, setChecking] = React.useState(false);
  const [checkCount, setCheckCount] = React.useState(0);
  const [contactLists, setContactLists] = React.useState<IContactListOption[] | null | undefined>(undefined);
  const [lookupColumns, setLookupColumns] = React.useState<ILookupColumnOption[] | null | undefined>(undefined);

  const [values, setValues] = React.useState<Values>({});
  const [dirty, setDirty] = React.useState<Flags>({});
  const [errors, setErrors] = React.useState<Errors>({});
  const [saving, setSaving] = React.useState(false);
  const [saveMessage, setSaveMessage] = React.useState<IMessage | undefined>(undefined);
  const [reloading, setReloading] = React.useState(false);
  const [setupBusy, setSetupBusy] = React.useState(false);
  const [setupMessage, setSetupMessage] = React.useState<IMessage | undefined>(undefined);
  const [confirmLeave, setConfirmLeave] = React.useState(false);

  const mounted = React.useRef(true);
  React.useEffect(() => () => { mounted.current = false; }, []);

  const loadContactLists = (): void => {
    if (!contactListSelectable) { return; }
    setupService.getContactLists().then(
      (lists: IContactListOption[]) => { if (mounted.current) { setContactLists(lists); } },
      () => { if (mounted.current) { setContactLists(null); } }
    );
  };
  React.useEffect(loadContactLists, [setupService]);

  const loadLookupColumns = (): void => {
    if (!contactColumnSelectable) { return; }
    setupService.getLookupColumns(page.listId).then(
      (columns: ILookupColumnOption[]) => { if (mounted.current) { setLookupColumns(columns); } },
      () => { if (mounted.current) { setLookupColumns(null); } }
    );
  };
  React.useEffect(loadLookupColumns, [setupService, page.listId]);

  const checkRun = React.useRef(0);
  React.useEffect(() => {
    if (!site) { return; }
    const run = ++checkRun.current;
    setChecking(true);
    setupService.check(site.settings, page.listId).then(
      (result: ISetupStatus) => { if (mounted.current && run === checkRun.current) { setStatus(result); setChecking(false); } },
      () => { if (mounted.current && run === checkRun.current) { setStatus(null); setChecking(false); } }
    );
  }, [site, checkCount, setupService, page.listId]);

  React.useEffect(() => {
    if (status && !tabChosen.current && !isSetupComplete(status)) {
      setTab('setup');
    }
    if (status !== undefined) {
      tabChosen.current = true;
    }
  }, [status]);

  const storedValue = (id: string): string => {
    const key = id === 'mailbox' ? base.mailbox.key : entryKey(settings.settingsList.keys, id);
    const stored = key && site && site.stored ? site.stored[key] : undefined;
    return typeof stored === 'string' ? stored : '';
  };
  const stored: Values = {};
  ids.forEach((id: string) => { stored[id] = storedValue(id); });
  const storedKey = JSON.stringify(stored);

  React.useEffect(() => {
    setValues((current: Values) => {
      const next = { ...current };
      ids.forEach((id: string) => {
        if (!dirty[id]) { next[id] = stored[id]; }
      });
      return next;
    });
  }, [storedKey, dirty]);

  React.useEffect(() => {
    if (reloading) {
      setReloading(false);
      setDirty({});
    }
  }, [site]);

  const listReady = status === null
    || (!!status && (status.settingsList === 'present' || status.settingsList === 'needsUpdate') && status.valueColumn === 'present');
  const listReadable = !!site && !!site.stored;
  const listEditable = canSaveSettings && listReadable && listReady;
  const editable = (id: string): boolean =>
    listEditable && !!(id === 'mailbox' ? base.mailbox.key : entryKey(settings.settingsList.keys, id));
  const anyDirty = ids.some((id: string) => !!dirty[id]);
  const busy = saving || setupBusy;

  const chooseTab = (next: Tab): void => {
    tabChosen.current = true;
    setTab(next);
  };

  const setField = (id: string, value: string): void => {
    setValues((current: Values) => ({ ...current, [id]: value }));
    setDirty((current: Flags) => ({ ...current, [id]: value !== stored[id] }));
    setErrors((current: Errors) => ({ ...current, [id]: undefined }));
    setSaveMessage(undefined);
    setConfirmLeave(false);
  };

  const validate = (id: string, value: string): string => {
    switch (parseId(id).field) {
      case 'contactList': return checkListName(value);
      case 'contactColumn': return checkColumnName(value);
      case 'categories': return checkCategories(value, policy);
      case 'mailSubject': return checkMailTemplate(value, policy.maxTextLength);
      case 'mailBody': return checkMailTemplate(value, policy.maxMailTemplateLength);
      default: return checkAddressField(value, policy);
    }
  };

  const save = (): void => {
    if (busy || !anyDirty) { return; }
    const changed = ids.filter((id: string) => dirty[id] && editable(id));
    const nextErrors: Errors = {};
    changed.forEach((id: string) => {
      const error = validate(id, values[id] || '');
      if (error) { nextErrors[id] = error; }
    });
    const firstError = changed.filter((id: string) => !!nextErrors[id])[0];
    if (firstError) {
      setErrors(nextErrors);
      chooseTab(tabOf(firstError));
      const errorLanguage = parseId(firstError).language;
      if (errorLanguage) { setLanguageCode(errorLanguage); }
      return;
    }

    const entries: IStoredValues = {};
    changed.forEach((id: string) => {
      entries[id === 'mailbox' ? base.mailbox.key : entryKey(settings.settingsList.keys, id)] = normalize(parseId(id).field, values[id] || '');
    });

    setSaving(true);
    setSaveMessage(undefined);
    setupService.saveValues(entries)
      .then(
        () => {
          if (!mounted.current) { return; }
          setSaving(false);
          setSaveMessage({ text: strings.settingsSaved, tone: 'ok' });
          setReloading(true);
          onChanged();
        },
        (error: unknown) => {
          if (!mounted.current) { return; }
          setSaving(false);
          setSaveMessage({ text: failureText(error, strings.settingsSaveFailed), tone: 'problem' });
        }
      );
  };

  const runSetup = (): void => {
    if (!status || !site || busy) { return; }
    setSetupBusy(true);
    setSetupMessage(undefined);
    setupService.setup(site.settings, page.listId, status).then(
      () => {
        if (!mounted.current) { return; }
        setSetupBusy(false);
        setSetupMessage({ text: strings.settingsSetupDone, tone: 'ok' });
        loadContactLists();
        loadLookupColumns();
        onChanged();
      },
      (error: unknown) => {
        if (!mounted.current) { return; }
        setSetupBusy(false);
        setSetupMessage({ text: failureText(error, strings.settingsSetupFailed), tone: 'problem' });
        setCheckCount((count: number) => count + 1);
      }
    );
  };

  const back = (): void => {
    if (anyDirty && !confirmLeave) {
      setConfirmLeave(true);
      return;
    }
    onBack();
  };

  const message = (msg: IMessage | undefined): React.ReactNode => msg && (
    <p className={msg.tone === 'ok' ? classNames.ok : classNames.problem} role={msg.tone === 'ok' ? 'status' : 'alert'}>{msg.text}</p>
  );

  const fieldProps = (id: string, multiline?: boolean): Partial<ITextFieldProps> => ({
    className: formClassNames.field,
    value: values[id] || '',
    onChange: (_e: unknown, value?: string) => setField(id, value || ''),
    onKeyDown: multiline ? undefined : (event: React.KeyboardEvent) => {
      if (event.key === 'Enter' && editable(id)) {
        event.preventDefault();
        save();
      }
    },
    errorMessage: errors[id],
    readOnly: !editable(id),
    maxLength: multiline ? undefined : policy.maxTextLength,
    disabled: busy,
    autoComplete: 'off'
  });

  const notices = (): React.ReactNode => {
    if (!listReadable) {
      return <MessageBar className={classNames.notice} messageBarType={MessageBarType.error}>{strings.settingsReadFailed}</MessageBar>;
    }
    if (!listReady) {
      return (
        <MessageBar className={classNames.notice} messageBarType={MessageBarType.warning}>
          {strings.settingsNotSetUp}{' '}
          <Link onClick={() => chooseTab('setup')}>{strings.settingsGoToSetup}</Link>
        </MessageBar>
      );
    }
    if (!canSaveSettings) {
      return <MessageBar className={classNames.notice}>{strings.settingsNoPermission}</MessageBar>;
    }
    return null;
  };

  const baseUrl = webUrl.replace(/\/+$/, '');
  const sameName = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();
  const urlName = (list: IListRef): string => {
    const match = /\/Lists\/([^/]+)$/i.exec(list.serverRelativeUrl);
    return match ? match[1] : '';
  };

  const defaultListName = base.contactList.url.substring('Lists/'.length);
  const contactListOptions = (): { options: IDropdownOption[]; selectedKey: string } => {
    const lists = (contactLists || []).filter((list: IContactListOption) => !sameName(list.name, defaultListName));
    const options: IDropdownOption[] = [
      { key: '', text: formatText(strings.settingsOptionDefault, base.contactList.title) },
      ...lists.map((list: IContactListOption) => ({
        key: list.name,
        text: sameName(list.title, list.name) ? list.title : `${list.title} (Lists/${list.name})`
      }))
    ];
    const current = (values.contactList || '').trim();
    if (!current || sameName(current, defaultListName)) {
      return { options, selectedKey: '' };
    }
    const match = lists.filter((list: IContactListOption) => sameName(list.name, current))[0];
    if (match) {
      return { options, selectedKey: match.name };
    }
    if (contactLists) {
      options.push({ key: current, text: formatText(strings.settingsOptionNotFound, current) });
    }
    return { options, selectedKey: current };
  };

  const lookupNotice = (): React.ReactNode => {
    const lookup = status ? status.lookupList : undefined;
    if (!lookup) { return null; }
    const selected = (values.contactList || '').trim() || defaultListName;
    if (sameName(urlName(lookup), selected)) { return null; }
    return <p className={classNames.warning}>{formatText(strings.settingsContactListLookup, lookup.title || strings.settingsSetupUnknownList)}</p>;
  };

  const readOnlyRow = (label: string, value: string): React.ReactNode => (
    <div className={classNames.row}>
      <div className={classNames.label}>{label}</div>
      <div className={classNames.value}>{value || strings.settingsStatusNotConfigured}</div>
    </div>
  );

  const contactListField = (): React.ReactNode => {
    if (!contactListSelectable) {
      const list = settings.contactList;
      const name = list.url.substring('Lists/'.length);
      return readOnlyRow(strings.settingsContactListLabel, !list.url ? '' : (sameName(list.title, name) ? list.title : `${list.title} (${list.url})`));
    }
    if (contactLists === null) {
      return (
        <TextField
          {...fieldProps('contactList')}
          label={strings.settingsContactListLabel}
          placeholder={defaultListName}
          description={strings.settingsContactListManual}
        />
      );
    }
    const { options, selectedKey } = contactListOptions();
    return (
      <div className={formClassNames.field}>
        <Dropdown
          label={strings.settingsContactListLabel}
          options={options}
          selectedKey={contactLists ? selectedKey : null}
          placeholder={strings.settingsContactListLoading}
          onChange={(_e: unknown, option?: IDropdownOption) => { if (option) { setField('contactList', String(option.key)); } }}
          errorMessage={errors.contactList}
          disabled={busy || !contactLists || !editable('contactList')}
        />
        <p className={classNames.description}>{strings.settingsContactListDescription}</p>
      </div>
    );
  };

  const defaultColumn = base.contactColumn.internalName;
  const columnText = (column: ILookupColumnOption): string =>
    (column.title === column.internalName ? column.title : `${column.title} (${column.internalName})`);
  const contactColumnOptions = (): { options: IDropdownOption[]; selectedKey: string } => {
    const all = lookupColumns || [];
    const existingDefault = all.filter((column: ILookupColumnOption) => column.internalName === defaultColumn)[0];
    const others = all.filter((column: ILookupColumnOption) => column.internalName !== defaultColumn);
    const options: IDropdownOption[] = [
      { key: '', text: formatText(strings.settingsOptionDefault, existingDefault ? columnText(existingDefault) : base.contactColumn.displayName) },
      ...others.map((column: ILookupColumnOption) => ({ key: column.internalName, text: columnText(column) }))
    ];
    const current = (values.contactColumn || '').trim();
    if (!current || current === defaultColumn) {
      return { options, selectedKey: '' };
    }
    if (lookupColumns && !others.some((column: ILookupColumnOption) => column.internalName === current)) {
      options.push({ key: current, text: formatText(strings.settingsOptionNotFound, current) });
    }
    return { options, selectedKey: current };
  };

  const contactColumnField = (): React.ReactNode => {
    if (!contactColumnSelectable) {
      const column = settings.contactColumn;
      return readOnlyRow(strings.settingsContactColumnLabel, !column.internalName ? ''
        : (column.displayName === column.internalName ? column.displayName : `${column.displayName} (${column.internalName})`));
    }
    if (lookupColumns === null) {
      return (
        <TextField
          {...fieldProps('contactColumn')}
          label={strings.settingsContactColumnLabel}
          placeholder={defaultColumn}
          description={strings.settingsContactColumnManual}
        />
      );
    }
    const { options, selectedKey } = contactColumnOptions();
    return (
      <div className={formClassNames.field}>
        <Dropdown
          label={strings.settingsContactColumnLabel}
          options={options}
          selectedKey={lookupColumns ? selectedKey : null}
          placeholder={strings.settingsContactColumnLoading}
          onChange={(_e: unknown, option?: IDropdownOption) => { if (option) { setField('contactColumn', String(option.key)); } }}
          errorMessage={errors.contactColumn}
          disabled={busy || !lookupColumns || !editable('contactColumn')}
        />
        <p className={classNames.description}>{strings.settingsContactColumnDescription}</p>
      </div>
    );
  };

  const pageContact = (): React.ReactNode => {
    const contact = defaults ? defaults.contact : undefined;
    const available = !!contact && ['ok', 'empty', 'invalid'].indexOf(contact.status) >= 0;
    const editFormUrl = `${baseUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${encodeURIComponent('{' + page.listId + '}')}&ID=${page.itemId}`;
    let content: React.ReactNode;
    let problem = false;
    if (!defaults) {
      content = strings.loadingRecipients;
    } else if (!contact || !available) {
      content = strings.settingsPageContactUnavailable;
    } else {
      problem = contact.status !== 'ok';
      content = (
        <>
          {contact.status === 'ok' ? contact.emails.join('; ')
            : (contact.status === 'invalid' ? strings.settingsPageContactInvalid : strings.settingsPageContactNone)}
          {page.canEditItems && <>{' · '}<Link href={editFormUrl} {...NEW_WINDOW}>{strings.settingsEditPageProperties}</Link></>}
        </>
      );
    }
    return (
      <div className={classNames.row}>
        <div className={classNames.label}>{strings.settingsPageContactLabel}</div>
        <div className={problem ? classNames.checkProblem : classNames.value}>{content}</div>
      </div>
    );
  };

  const recipientsTab = (): React.ReactNode => (
    <>
      {notices()}
      <div className={formClassNames.columns}>
        <section className={formClassNames.column} aria-labelledby="spfx-feedback-settings-contact">
          <h3 id="spfx-feedback-settings-contact" className={formClassNames.sectionTitle}>{strings.settingsContactHeading}</h3>
          {contactListField()}
          {lookupNotice()}
          {contactColumnField()}
          {(!contactListSelectable || !contactColumnSelectable) && <p className={classNames.description}>{strings.settingsContactFixed}</p>}
          {pageContact()}
        </section>

        <section className={formClassNames.column} aria-labelledby="spfx-feedback-settings-mailbox">
          <h3 id="spfx-feedback-settings-mailbox" className={formClassNames.sectionTitle}>{strings.settingsMailboxHeading}</h3>
          {base.mailbox.key ? (
            <TextField
              {...fieldProps('mailbox')}
              label={strings.settingsMailboxLabel}
              placeholder={strings.mailboxPlaceholder}
              description={formatText(strings.settingsMailboxDescriptionList, base.settingsList.title)}
              maxLength={policy.maxTextLength}
              inputMode="email"
            />
          ) : message({ text: strings.settingsMailboxNotConfigured, tone: 'problem' })}
        </section>
      </div>
    </>
  );

  const languageId = (field: LanguageField): string => fieldId(field, language.code);

  const defaultFor = (field: LanguageField, entry: ILanguageSettings): string =>
    (field === 'categories' ? entry.categories.join('\n') : (field === 'mailSubject' ? entry.subject : entry.body));

  const isCustomized = (field: LanguageField, entry: ILanguageSettings): boolean => {
    const value = normalize(field, values[fieldId(field, entry.code)] || '');
    return !!value && value !== normalize(field, defaultFor(field, entry));
  };

  const languageHasErrors =(code: string): boolean => LANGUAGE_FIELDS.some((field: LanguageField) => !!errors[fieldId(field, code)]);

  const defaultLinks = (fields: LanguageField[]): React.ReactNode => {
    if (!listEditable) { return null; }
    const empty = fields.every((field: LanguageField) => !(values[languageId(field)] || '').trim());
    const customized = fields.some((field: LanguageField) => isCustomized(field, language));
    if (!empty && !customized) { return null; }
    const restore = (): void => fields.forEach((field: LanguageField) => setField(languageId(field), defaultFor(field, language)));
    return (
      <p className={classNames.links}>
        <Link onClick={restore} disabled={busy}>{empty ? strings.settingsInsertDefault : strings.settingsUseDefault}</Link>
      </p>
    );
  };

  const languagePicker = (): React.ReactNode => {
    if (languages.length < 2) { return null; }
    const withErrors = languages.filter((entry: ILanguageSettings) => entry.code !== language.code && languageHasErrors(entry.code));
    return (
      <Dropdown
        className={classNames.language}
        label={strings.settingsLanguageLabel}
        selectedKey={language.code}
        options={languages.map((entry: ILanguageSettings) => ({
          key: entry.code,
          text: LANGUAGE_FIELDS.some((field: LanguageField) => isCustomized(field, entry))
            ? formatText(strings.settingsLanguageCustomized, entry.label)
            : entry.label
        }))}
        onChange={(_e: unknown, option?: IDropdownOption) => { if (option) { setLanguageCode(String(option.key)); } }}
        errorMessage={withErrors.length > 0
          ? formatText(strings.settingsLanguageErrors, withErrors.map((entry: ILanguageSettings) => entry.label).join(', '))
          : undefined}
        disabled={busy}
      />
    );
  };

  const labelOf = (code: string): string =>
    (languages.filter((entry: ILanguageSettings) => entry.code === code)[0] || languages[0]).label;

  const mailLanguageNote = (): React.ReactNode => (
    <p className={classNames.languageNote}>
      {languages.length < 2
        ? formatText(strings.settingsMailLanguageSingle, labelOf(base.language))
        : formatText(strings.settingsMailLanguage, labelOf(base.language), labelOf(base.defaultLanguage))}
    </p>
  );

  const placeholders = MAIL_PLACEHOLDERS.map((name: string) => `{{${name}}}`).join(' ');

  const textsTab = (): React.ReactNode => (
    <>
      {notices()}
      {languagePicker()}
      {mailLanguageNote()}
      <div className={formClassNames.columns}>
        <section className={formClassNames.column} aria-labelledby="spfx-feedback-settings-form">
          <h3 id="spfx-feedback-settings-form" className={formClassNames.sectionTitle}>{strings.settingsFormHeading}</h3>
          <TextField
            key={languageId('categories')}
            {...fieldProps(languageId('categories'), true)}
            label={strings.settingsCategoriesLabel}
            placeholder={language.categories.join('\n')}
            description={strings.settingsCategoriesDescription}
            multiline
            rows={10}
          />
          {defaultLinks(['categories'])}
        </section>

        <section className={formClassNames.column} aria-labelledby="spfx-feedback-settings-mail">
          <h3 id="spfx-feedback-settings-mail" className={formClassNames.sectionTitle}>{strings.settingsMailHeading}</h3>
          <TextField
            key={languageId('mailSubject')}
            {...fieldProps(languageId('mailSubject'))}
            label={strings.settingsMailSubjectLabel}
            placeholder={language.subject}
            maxLength={policy.maxTextLength}
          />
          <TextField
            key={languageId('mailBody')}
            {...fieldProps(languageId('mailBody'), true)}
            label={strings.settingsMailBodyLabel}
            placeholder={strings.settingsMailBodyPlaceholder}
            description={formatText(strings.settingsMailDescription, placeholders)}
            maxLength={policy.maxMailTemplateLength}
            multiline
            rows={8}
          />
          {defaultLinks(['mailSubject', 'mailBody'])}
        </section>
      </div>
    </>
  );

  const statusText = (part: SetupPartStatus): string => {
    switch (part) {
      case 'present': return strings.settingsStatusPresent;
      case 'missing': return strings.settingsStatusMissing;
      case 'needsUpdate': return strings.settingsStatusVisible;
      case 'broken': return strings.settingsStatusBroken;
      case 'wrongType': return strings.settingsStatusWrongType;
      default: return strings.settingsStatusNotConfigured;
    }
  };

  const checkItem = (key: string, label: string, part: SetupPartStatus, detail?: string, text?: string): React.ReactNode => {
    const fixable = part === 'missing' || part === 'needsUpdate' || part === 'broken';
    const icon = part === 'present' ? 'CompletedSolid' : (fixable ? 'Warning' : 'ErrorBadge');
    const iconClass = part === 'present' ? classNames.iconOk : (fixable ? classNames.iconWarning : classNames.iconError);
    return (
      <li key={key} className={classNames.checkItem}>
        <Icon iconName={icon} className={`${classNames.checkIcon} ${iconClass}`} aria-hidden="true" />
        <div>
          <p className={classNames.checkLabel}>{label}</p>
          <p className={part === 'present' || fixable ? classNames.checkDetail : classNames.checkProblem}>
            {text || statusText(part)}{detail ? ` · ${detail}` : ''}
          </p>
        </div>
      </li>
    );
  };

  const settingsListItem = (current: ISetupStatus): React.ReactNode => {
    const label = formatText(strings.settingsSetupSettingsList, base.settingsList.title);
    if (current.settingsList !== 'present') {
      return checkItem('settingsList', label, current.settingsList);
    }
    const incomplete = [current.valueColumn, current.entries].filter((part: SetupPartStatus) => part !== 'present');
    if (incomplete.length === 0) {
      return checkItem('settingsList', label, 'present');
    }
    const part = incomplete.filter((p: SetupPartStatus) => p !== 'missing')[0] || 'needsUpdate';
    return checkItem('settingsList', label, part, undefined, part === 'needsUpdate' ? strings.settingsStatusIncomplete : undefined);
  };

  const checklist = (current: ISetupStatus): React.ReactNode => {
    const lookup = current.lookupList;
    const contactTitle = lookup ? (lookup.title || strings.settingsSetupUnknownList) : settings.contactList.title;
    const lookupElsewhere = !!lookup && !sameName(urlName(lookup), settings.contactList.url.substring('Lists/'.length));
    return (
      <ul className={classNames.checklist}>
        {settingsListItem(current)}
        {checkItem('contactList', formatText(strings.settingsSetupContactList, contactTitle), current.contactList,
          lookupElsewhere ? formatText(strings.settingsSetupLookupTarget, settings.contactList.title) : undefined)}
        {checkItem('emailColumn', formatText(strings.settingsSetupEmailColumn, settings.contactEmailColumn.displayName), current.emailColumn)}
        {checkItem('pageColumn', formatText(strings.settingsSetupPageColumn, settings.contactColumn.displayName), current.pageColumn)}
      </ul>
    );
  };

  const setupTab = (): React.ReactNode => {
    if (status === null) {
      return (
        <>
          {message({ text: strings.settingsSetupCheckFailed, tone: 'problem' })}
          <div className={classNames.actions}>
            <DefaultButton iconProps={{ iconName: 'Refresh' }} text={strings.settingsSetupRefresh}
              onClick={() => setCheckCount((count: number) => count + 1)} disabled={busy || checking} />
          </div>
        </>
      );
    }
    if (!status) { return null; }
    const complete = isSetupComplete(status);
    const fixable = canSetUp(status);
    const contactDirty = !!dirty.contactList || !!dirty.contactColumn;
    const hint = !fixable ? (complete ? undefined : strings.settingsSetupNothingToDo)
      : (!page.canManageList ? strings.settingsSetupNeedsPermission : (contactDirty ? strings.settingsSetupSaveFirst : undefined));
    return (
      <>
        <p className={classNames.intro}>{strings.settingsSetupIntro}</p>
        {checklist(status)}
        {complete && !setupMessage && message({ text: strings.settingsSetupComplete, tone: 'ok' })}
        {hint && <p className={classNames.description}>{hint}</p>}
        <div className={classNames.actions}>
          {fixable && (
            <DefaultButton
              iconProps={{ iconName: 'Repair' }}
              text={strings.settingsSetupRun}
              onClick={runSetup}
              disabled={busy || checking || !page.canManageList || contactDirty}
            >
              {setupBusy && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
            </DefaultButton>
          )}
          <DefaultButton
            iconProps={{ iconName: 'Refresh' }}
            text={strings.settingsSetupRefresh}
            onClick={() => setCheckCount((count: number) => count + 1)}
            disabled={busy || checking}
          >
            {checking && !setupBusy && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
          </DefaultButton>
        </div>
        {message(setupMessage)}
      </>
    );
  };

  const loading = !site || status === undefined;
  const tabHasErrors = (id: Tab): boolean => ids.some((fieldKey: string) => !!errors[fieldKey] && tabOf(fieldKey) === id);
  const tabWarning = (id: Tab): boolean => (id === 'setup' ? !!status && !isSetupComplete(status) : tabHasErrors(id));
  const tabId = (id: Tab): string => `spfx-feedback-settings-tab-${id}`;
  const tabItem = (id: Tab, headerText: string): React.ReactElement => (
    <PivotItem
      itemKey={id}
      headerText={headerText}
      itemIcon={tabWarning(id) ? 'Warning' : undefined}
      ariaLabel={tabWarning(id) ? formatText(strings.settingsTabWarning, headerText) : undefined}
    />
  );

  return (
    <>
      <Pivot
        className={classNames.tabs}
        headersOnly
        selectedKey={tab}
        getTabId={(itemKey: string) => tabId(itemKey as Tab)}
        onLinkClick={(item?: PivotItem) => { if (item && item.props.itemKey) { chooseTab(item.props.itemKey as Tab); } }}
      >
        {tabItem('recipients', strings.settingsTabRecipients)}
        {tabItem('texts', strings.settingsTabTexts)}
        {tabItem('setup', strings.settingsTabSetup)}
      </Pivot>

      {loading ? (
        <Spinner label={strings.loadingRecipients} className={classNames.loading} />
      ) : (
        <div className={classNames.panel} role="tabpanel" aria-labelledby={tabId(tab)}>
          {tab === 'recipients' && recipientsTab()}
          {tab === 'texts' && textsTab()}
          {tab === 'setup' && setupTab()}
          {message(saveMessage)}
        </div>
      )}

      {confirmLeave && (
        <MessageBar
          className={classNames.notice}
          messageBarType={MessageBarType.warning}
          actions={
            <div>
              <DefaultButton text={strings.settingsDiscard} onClick={onBack} />
              <DefaultButton text={strings.settingsKeepEditing} onClick={() => setConfirmLeave(false)} />
            </div>
          }
        >
          {strings.settingsUnsaved}
        </MessageBar>
      )}

      <DialogFooter>
        {canSaveSettings && (
          <PrimaryButton text={strings.settingsSave} onClick={save} disabled={loading || busy || !anyDirty}>
            {saving && <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginLeft: 8 } }} />}
          </PrimaryButton>
        )}
        <DefaultButton text={strings.settingsBack} onClick={back} disabled={busy} />
      </DialogFooter>
    </>
  );
};
