import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { encodePropertyName, GUID, listUrl, NOMETADATA } from './RecipientService';
import { IFeedbackSettings, IListRef, ISetupStatus, IStoredValues, SetupPartStatus } from '../models/IFeedback';
import { logError, logInfo, logVerbose } from '../utils/logging';
import { defaultSettingEntries, FIELD_NAME, LIST_NAME } from '../utils/sanitize';
import { IUntyped, odataEntity, odataItems } from '../utils/untyped';

/** A list that can be chosen as contact list: name in its URL ("Lists/<name>") and title. */
export interface IContactListOption {
  name: string;
  title: string;
}

/** A lookup column of the Site Pages library that can be chosen as contact column. */
export interface ILookupColumnOption {
  internalName: string;
  title: string;
}

/** SP.FieldType.Lookup (single and multiple values). */
const FIELD_TYPE_LOOKUP = 7;

/** A custom list of the site (getGenericLists). */
interface IListInfo extends IListRef {
  hidden: boolean;
}

/** Entries of the settings list: item type (for creating entries) and ID per title. */
interface IEntries {
  type: string;
  ids: { [key: string]: number };
}

/** Schema of a field, undefined if the list has no field with this internal name. */
interface IFieldSchema {
  type: string;
  lookupList: string;
}

/** Base template of a custom list ("Generic List"). */
const GENERIC_LIST_TEMPLATE = 100;

const LOOKUP_TYPES = ['Lookup', 'LookupMulti'];

/** SP.AddFieldOptions */
const ADD_TO_DEFAULT_CONTENT_TYPE = 1;
const ADD_TO_ALL_CONTENT_TYPES = 4;
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;

/** TypeId of SP.ClientContext, used for "ClientContext.Current" in CSOM requests. */
const CLIENT_CONTEXT_TYPE_ID = '{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}';

/** Verbose JSON (with __metadata); the empty OData-Version header suppresses SPHttpClient's "4.0". */
const VERBOSE_HEADERS = { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose', 'OData-Version': '' };

/** Error of a setup action; accessDenied = missing permission (or custom scripts blocked in SharePoint Online). */
export class SetupError extends Error {
  constructor(message: string, public readonly accessDenied: boolean) {
    super(message);
    // ES5 target: restore the prototype chain, otherwise instanceof fails.
    Object.setPrototypeOf(this, SetupError.prototype);
  }
}

/** True if nothing is missing (parts that cannot be fixed automatically count as incomplete). */
export function isSetupComplete(status: ISetupStatus): boolean {
  return [status.settingsList, status.valueColumn, status.entries, status.contactList, status.emailColumn, status.pageColumn]
    .every((part: SetupPartStatus) => part === 'present');
}

/** True if "Set up" can create or fix at least one part. */
export function canSetUp(status: ISetupStatus): boolean {
  const settingsMissing = [status.settingsList, status.valueColumn, status.entries]
    .some((part: SetupPartStatus) => part === 'missing' || part === 'needsUpdate');
  const contactMissing = status.pageColumn !== 'wrongType' && status.pageColumn !== 'notConfigured'
    && [status.contactList, status.emailColumn, status.pageColumn].some((part: SetupPartStatus) => part === 'missing' || part === 'broken');
  return settingsMissing || contactMissing;
}

/**
 * Everything the settings view (gear icon) does in SharePoint - and only the settings view:
 *  - check: which lists and columns exist
 *  - setup: creates what is missing - settings list with value column, contact list with e-mail column and
 *    the lookup column in the Site Pages library. Every step checks first whether its part already exists,
 *    so the setup can simply be repeated after a partial failure.
 *  - saveValues: entries of the settings list (the list must have been set up)
 *  - shared mailbox in the property bag of the web (config mailbox.storage "propertyBag")
 *
 * SharePoint checks the permissions on the server.
 */
export class SetupService {
  private readonly webUrl: string;
  private readonly api: string;

  constructor(
    private readonly httpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly webServerRelativeUrl: string,
    /** Code configuration: settings list and mailbox storage (not changeable in the settings list). */
    private readonly base: IFeedbackSettings
  ) {
    this.webUrl = webAbsoluteUrl.replace(/\/+$/, '');
    this.api = this.webUrl + '/_api';
  }

  /**
   * Checks the parts of the setup with the current settings of the site (code configuration and settings list).
   * Custom lists and the lookup column in parallel; the target of the lookup column if it is no custom list
   * (404 = deleted -> status broken); then value column, e-mail column and entries in parallel.
   * Lists that do not exist count as missing, they never make the check fail.
   */
  public check(settings: IFeedbackSettings, sitePagesListId: string): Promise<ISetupStatus> {
    const column = settings.contactColumn.internalName;
    const settingsUrl = this.base.settingsList.url;
    return Promise.all([
      this.getGenericLists(),
      column ? this.readField(sitePagesListId, column) : Promise.resolve(undefined)
    ]).then(([lists, pageField]: [IListInfo[], IFieldSchema | undefined]) => {
      const settingsList = settingsUrl ? this.byUrl(lists, settingsUrl) : undefined;
      const contactList = settings.contactList.url ? this.byUrl(lists, settings.contactList.url) : undefined;
      const isLookup = !!pageField && LOOKUP_TYPES.indexOf(pageField.type) >= 0;
      const lookupMatch = pageField && isLookup ? GUID.exec(pageField.lookupList) : null;
      // The target of the lookup column may be any list - or one that was deleted (-> broken).
      return (lookupMatch ? this.findListById(lists, lookupMatch[1]) : Promise.resolve(undefined))
        .then((lookupList: IListInfo | undefined) => this.checkParts(settings, pageField, isLookup && !!lookupMatch,
          { settings: settingsList, contact: contactList, lookup: lookupList }));
    });
  }

  /** Second part of check(): value column, e-mail column and entries, then the status of every part. */
  private checkParts(
    settings: IFeedbackSettings,
    pageField: IFieldSchema | undefined,
    isLookup: boolean,
    found: { settings?: IListInfo; contact?: IListInfo; lookup?: IListInfo }
  ): Promise<ISetupStatus> {
    const column = settings.contactColumn.internalName;
    const emailColumn = settings.contactEmailColumn.internalName;
    const valueColumn = this.base.settingsList.valueColumn.internalName;
    const settingsUrl = this.base.settingsList.url;
    const { settings: settingsList, contact: contactList, lookup: lookupList } = found;
    /** The lookup column exists, but its list was deleted: "Set up" creates the column anew. */
    const broken = isLookup && !lookupList;
    // The form reads the e-mail column of the list the lookup column points to (a broken one: the contact list).
    const emailList = lookupList || contactList;
    const expected = Object.keys(defaultSettingEntries(this.base));
    return Promise.all([
      settingsList && valueColumn ? this.readField(settingsList.id, valueColumn, true) : Promise.resolve(undefined),
      emailList && emailColumn ? this.readField(emailList.id, emailColumn, true) : Promise.resolve(undefined),
      settingsList ? this.readEntries(settingsList.id) : Promise.resolve(undefined)
    ]).then(([valueField, emailField, entries]: [IFieldSchema | undefined, IFieldSchema | undefined, IEntries | undefined]): ISetupStatus => {
      const missingEntries = expected.filter((key: string) => !entries || !Object.prototype.hasOwnProperty.call(entries.ids, key));
      const status: ISetupStatus = {
        settingsList: !settingsUrl ? 'notConfigured' : (!settingsList ? 'missing' : (settingsList.hidden ? 'present' : 'needsUpdate')),
        valueColumn: !settingsUrl ? 'notConfigured' : (valueField ? 'present' : 'missing'),
        entries: !settingsUrl ? 'notConfigured' : (missingEntries.length > 0 ? 'missing' : 'present'),
        missingEntries,
        // With a lookup column its target is the contact list that counts (setup creates no other list then).
        contactList: lookupList ? 'present' : (!settings.contactList.url ? 'notConfigured' : (contactList ? 'present' : 'missing')),
        emailColumn: !emailColumn ? 'notConfigured' : (emailField ? 'present' : 'missing'),
        pageColumn: !column ? 'notConfigured' : (!pageField ? 'missing' : (lookupList ? 'present' : (broken ? 'broken' : 'wrongType'))),
        lists: { settings: settingsList, contact: contactList, lookup: lookupList }
      };
      logVerbose('Setup status', status);
      return status;
    });
  }

  /** A list by its ID: from the custom lists, otherwise read (other templates); undefined if it does not exist. */
  private findListById(lists: IListInfo[], id: string): Promise<IListInfo | undefined> {
    const known = lists.filter((list: IListInfo) => list.id.toLowerCase() === id.toLowerCase())[0];
    if (known) {
      return Promise.resolve(known);
    }
    const url = `${listUrl(this.api, id)}?$select=Id,Title,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    return this.get(url, 'List could not be read', true)
      .then((json: unknown) => (json === undefined ? undefined : toListInfo(odataEntity(json))));
  }

  /**
   * Creates everything that is missing, with the current settings of the site:
   *  1. settings list (found by its URL; created if missing, hidden in the site contents), its value column and
   *     one entry per setting with the default of the code (existing entries are left unchanged)
   *  2. contact list - unless the lookup column already exists, then its target list is used
   *  3. e-mail column in that list
   *  4. lookup column in Site Pages (if missing), added to all content types -> visible in the page properties
   * A column with the name of the lookup column but another type is left alone (status wrongType).
   * A lookup column whose list was deleted (status broken) is deleted first and created anew for the contact list:
   * the target of a lookup column cannot be changed, and its values point to nothing anymore.
   */
  public setup(settings: IFeedbackSettings, sitePagesListId: string, status: ISetupStatus): Promise<void> {
    const settingsPart = this.base.settingsList.url
      ? this.ensureSettingsList().then((listId: string) => this.ensureEntries(listId))
      : Promise.resolve(undefined);
    const contactPossible = !!settings.contactColumn.internalName && !!settings.contactList.url
      && status.pageColumn !== 'wrongType';
    const removeBroken = (): Promise<unknown> => (status.pageColumn === 'broken'
      ? this.deleteField(sitePagesListId, settings.contactColumn.internalName)
      : Promise.resolve(undefined));
    return settingsPart
      .then(() => (contactPossible
        ? removeBroken().then(() => this.setupContact(settings, sitePagesListId, status.pageColumn === 'present' ? status.lists.lookup : undefined))
        : undefined))
      .then(() => logInfo('Setup completed'));
  }

  /**
   * Custom lists of the site that can hold the contact persons: not hidden, directly below "Lists/",
   * with a URL name allowed in the settings (LIST_NAME). The settings list itself is left out. Sorted by title.
   */
  public getContactLists(): Promise<IContactListOption[]> {
    const prefix = (this.webServerRelativeUrl.replace(/\/+$/, '') + '/Lists/').toLowerCase();
    const settingsUrl = (this.base.settingsList.url || '').toLowerCase();
    return this.getGenericLists()
      .then((lists: IListInfo[]) => lists
        .map((list: IListInfo): IContactListOption | undefined => {
          const url = list.serverRelativeUrl;
          if (list.hidden || url.toLowerCase().indexOf(prefix) !== 0) {
            return undefined;
          }
          const name = url.substring(prefix.length);
          if (!LIST_NAME.test(name) || ('lists/' + name).toLowerCase() === settingsUrl) {
            return undefined;
          }
          return { name, title: list.title || name };
        })
        .filter((option: IContactListOption | undefined): option is IContactListOption => !!option)
        .sort((a: IContactListOption, b: IContactListOption) => a.title.localeCompare(b.title)));
  }

  /**
   * Lookup columns of the Site Pages library that can hold the contact person: visible, not read-only (no system
   * columns), with an internal name allowed in the settings (FIELD_NAME). Sorted by title.
   */
  public getLookupColumns(sitePagesListId: string): Promise<ILookupColumnOption[]> {
    const filter = encodeURIComponent(`FieldTypeKind eq ${FIELD_TYPE_LOOKUP} and Hidden eq false and ReadOnlyField eq false`);
    return this.get(`${listUrl(this.api, sitePagesListId)}/fields?$select=InternalName,Title&$filter=${filter}`, 'Columns could not be read')
      .then((json: unknown) => odataItems(json)
        .map((field: IUntyped): ILookupColumnOption => {
          const internalName = String(field.InternalName || '');
          return { internalName, title: typeof field.Title === 'string' && field.Title.trim() ? field.Title.trim() : internalName };
        })
        .filter((option: ILookupColumnOption) => FIELD_NAME.test(option.internalName) && option.internalName.charAt(0) !== '_')
        .sort((a: ILookupColumnOption, b: ILookupColumnOption) => a.title.localeCompare(b.title)));
  }

  /** Raw value of the shared mailbox in the property bag (storage "propertyBag"); '' if the key does not exist. */
  public getPropertyBagMailbox(): Promise<string> {
    const name = encodePropertyName(this.base.mailbox.key);
    return this.get(`${this.api}/web/AllProperties?$select=${name}`, 'Shared mailbox could not be read')
      .then((json: unknown) => {
        const value = odataEntity(json)[name];
        return typeof value === 'string' ? value : '';
      });
  }

  /** Stores the shared mailbox where the configuration says (settings list or property bag). */
  public setMailbox(value: string): Promise<void> {
    const { storage, key } = this.base.mailbox;
    return storage === 'list' ? this.saveValues({ [key]: value }) : this.setMailboxInPropertyBag(value);
  }

  /**
   * Stores entries in the settings list: updates the entry whose title is the key (the first one, as
   * SettingsService reads it), or creates it. An empty value clears an existing entry (= default of the code)
   * and creates none. The list is not created here - that is the job of setup().
   * Plain REST: works with "Design" permission and without custom scripts.
   */
  public saveValues(values: IStoredValues): Promise<void> {
    const keys = Object.keys(values).filter((key: string) => !!key);
    const { url } = this.base.settingsList;
    if (keys.length === 0) {
      return Promise.resolve();
    }
    if (!url) {
      return Promise.reject(new SetupError('Settings list not configured', false));
    }
    const column = this.base.settingsList.valueColumn.internalName;
    return this.findList(url)
      .then((existingList: IListInfo | undefined) => {
        if (!existingList) {
          throw new SetupError('Settings list not set up', false);
        }
        const list = listUrl(this.api, existingList.id);
        return this.readEntries(existingList.id).then((entries: IEntries) =>
          // One after the other: only a few entries, and a failure stops at a defined point.
          keys.reduce((chain: Promise<unknown>, key: string) => chain.then(() => {
            const body: { [name: string]: unknown } = { __metadata: { type: entries.type }, Title: key };
            body[column] = values[key];
            if (Object.prototype.hasOwnProperty.call(entries.ids, key)) {
              return this.post(`${list}/items(${entries.ids[key]})`, body, 'Setting could not be stored', { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
            }
            return values[key] ? this.post(`${list}/items`, body, 'Setting could not be stored') : undefined;
          }), Promise.resolve<unknown>(undefined)));
      })
      .then(() => logInfo('Settings stored', { list: url, keys }));
  }

  /** Settings list (hidden in the site contents) and its value column (multiple lines of text, in the default view). */
  private ensureSettingsList(): Promise<string> {
    const { url, title, valueColumn } = this.base.settingsList;
    const column = valueColumn.internalName;
    const schemaXml = `<Field Type="Note" Name="${column}" StaticName="${column}"`
      + ` DisplayName="${escapeXml(valueColumn.displayName)}" NumLines="6" RichText="FALSE" />`;
    return this.ensureList(url, title, true)
      .then((listId: string) => this.ensureField(listId, column, schemaXml,
        ADD_TO_DEFAULT_CONTENT_TYPE | ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW).then(() => listId));
  }

  /**
   * Creates every entry of the settings list that does not exist yet, with the default of the code configuration
   * (also empty ones, e.g. the shared mailbox, so that every setting is visible in the list). Existing entries keep their value.
   */
  private ensureEntries(listId: string): Promise<void> {
    const defaults = defaultSettingEntries(this.base);
    const column = this.base.settingsList.valueColumn.internalName;
    const list = listUrl(this.api, listId);
    return this.readEntries(listId).then((entries: IEntries) => {
      const missing = Object.keys(defaults).filter((key: string) => !Object.prototype.hasOwnProperty.call(entries.ids, key));
      return missing.reduce((chain: Promise<unknown>, key: string) => chain.then(() => {
        const body: { [name: string]: unknown } = { __metadata: { type: entries.type }, Title: key };
        body[column] = defaults[key];
        return this.post(`${list}/items`, body, 'Setting could not be created');
      }), Promise.resolve<unknown>(undefined)).then(() => {
        if (missing.length > 0) {
          logInfo('Settings entries created', { keys: missing });
        }
      });
    });
  }

  /** Item type of the settings list and the ID of each entry by its title (the first one, as SettingsService reads it). */
  private readEntries(listId: string): Promise<IEntries> {
    const list = listUrl(this.api, listId);
    return Promise.all([
      this.get(`${list}?$select=ListItemEntityTypeFullName`, 'List could not be read'),
      this.get(`${list}/items?$select=Id,Title&$orderby=Id&$top=500`, 'Settings could not be read')
    ]).then(([info, items]: [unknown, unknown]): IEntries => {
      const ids: { [key: string]: number } = {};
      odataItems(items).forEach((item: IUntyped) => {
        const key = typeof item.Title === 'string' ? item.Title.trim() : '';
        if (key && typeof item.Id === 'number' && !Object.prototype.hasOwnProperty.call(ids, key)) {
          ids[key] = item.Id;
        }
      });
      return { type: String(odataEntity(info).ListItemEntityTypeFullName || ''), ids };
    });
  }

  /** Contact list (or the existing target of the lookup column), its e-mail column, then the lookup column. */
  private setupContact(settings: IFeedbackSettings, sitePagesListId: string, existingLookupList: IListRef | undefined): Promise<void> {
    const listIdPromise = existingLookupList && GUID.test(existingLookupList.id)
      ? Promise.resolve(existingLookupList.id)
      : this.ensureList(settings.contactList.url, settings.contactList.title);
    return listIdPromise
      .then((contactListId: string) => this.ensureEmailColumn(settings, contactListId).then(() => contactListId))
      .then((contactListId: string) => this.ensurePageColumn(settings, sitePagesListId, contactListId));
  }

  /**
   * Property bag of the web. REST cannot write the property bag, therefore a CSOM request
   * (client.svc/ProcessQuery) sets web.AllProperties[key] and calls web.Update().
   */
  private setMailboxInPropertyBag(value: string): Promise<void> {
    const key = this.base.mailbox.key;
    const xml = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0"'
      + ' LibraryVersion="16.0.0.0" ApplicationName="spfx-feedback">'
      + '<Actions>'
      + '<Method Name="SetFieldValue" Id="1" ObjectPathId="4"><Parameters>'
      + `<Parameter Type="String">${escapeXml(key)}</Parameter>`
      + `<Parameter Type="String">${escapeXml(value)}</Parameter>`
      + '</Parameters></Method>'
      + '<Method Name="Update" Id="2" ObjectPathId="2" />'
      + '</Actions>'
      + '<ObjectPaths>'
      + `<StaticProperty Id="0" TypeId="${CLIENT_CONTEXT_TYPE_ID}" Name="Current" />`
      + '<Property Id="2" ParentId="0" Name="Web" />'
      + '<Property Id="4" ParentId="2" Name="AllProperties" />'
      + '</ObjectPaths>'
      + '</Request>';
    const url = this.webUrl + '/_vti_bin/client.svc/ProcessQuery';
    logVerbose('POST ' + url, { key });
    return this.httpClient.post(url, SPHttpClient.configurations.v1, {
      // client.svc is no REST URL: SPHttpClient needs the web URL for the request digest.
      webUrl: this.webUrl,
      headers: { 'Content-Type': 'text/xml', 'Accept': 'application/json', 'OData-Version': '' },
      body: xml
    })
      .then((response: SPHttpClientResponse) => this.checked(response, 'Mailbox could not be stored'))
      .then((json: unknown) => {
        // CSOM answers with HTTP 200 and reports errors in the first array element.
        const info = Array.isArray(json) && json[0] && typeof json[0] === 'object' ? (json[0] as IUntyped).ErrorInfo : undefined;
        if (info && typeof info === 'object') {
          const error = info as IUntyped;
          const denied = /UnauthorizedAccess|AccessDenied/i.test(String(error.ErrorTypeName || ''));
          logError('Mailbox could not be stored', { error: error.ErrorMessage, type: error.ErrorTypeName });
          throw new SetupError(String(error.ErrorMessage || 'CSOM error'), denied);
        }
        logInfo('Shared mailbox stored', { key });
      });
  }

  /** Custom lists of the site (template 100) with ID, title, URL and hidden flag. */
  private getGenericLists(): Promise<IListInfo[]> {
    const filter = encodeURIComponent(`BaseTemplate eq ${GENERIC_LIST_TEMPLATE}`);
    return this.get(`${this.api}/web/lists?$select=Id,Title,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=${filter}`, 'Lists could not be read')
      .then((json: unknown) => odataItems(json)
        .map(toListInfo)
        .filter((list: IListInfo | undefined): list is IListInfo => !!list));
  }

  /** The list with this URL ("Lists/<name>"), compared without case. */
  private byUrl(lists: IListInfo[], url: string): IListInfo | undefined {
    const serverRelativeUrl = (this.webServerRelativeUrl.replace(/\/+$/, '') + '/' + url).toLowerCase();
    return lists.filter((list: IListInfo) => list.serverRelativeUrl.toLowerCase() === serverRelativeUrl)[0];
  }

  private findList(url: string): Promise<IListInfo | undefined> {
    return this.getGenericLists().then((lists: IListInfo[]) => this.byUrl(lists, url));
  }

  /**
   * ID of the list with this URL ("Lists/<name>"); the list is created if it does not exist.
   * hidden = true: the list is hidden in the site contents (an existing visible list is hidden, too).
   */
  private ensureList(url: string, title: string, hidden?: boolean): Promise<string> {
    return this.findList(url).then((existing: IListInfo | undefined) => {
      if (!existing) {
        return this.createList(url.substr('Lists/'.length), title, !!hidden);
      }
      logVerbose('List exists', { url });
      if (!hidden || existing.hidden) {
        return existing.id;
      }
      return this.updateList(existing.id, { Hidden: true }, 'List could not be hidden')
        .then(() => { logInfo('List hidden', { url }); return existing.id; });
    });
  }

  private updateList(listId: string, properties: { [name: string]: unknown }, message: string): Promise<unknown> {
    return this.post(listUrl(this.api, listId), { __metadata: { type: 'SP.List' }, ...properties }, message,
      { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  }

  /** The URL of a list comes from the title at creation; a different title and the hidden flag are set afterwards. */
  private createList(urlName: string, title: string, hidden: boolean): Promise<string> {
    const body = {
      __metadata: { type: 'SP.List' },
      BaseTemplate: GENERIC_LIST_TEMPLATE,
      Title: urlName,
      Description: '',
      AllowContentTypes: false
    };
    return this.post(`${this.api}/web/lists`, body, 'List could not be created')
      .then((json: unknown) => {
        const id = String(odataEntity(json).Id || '');
        if (!GUID.test(id)) {
          throw new SetupError('List created, but no ID returned', false);
        }
        logInfo('List created', { url: 'Lists/' + urlName, hidden });
        const changes: { [name: string]: unknown } = {};
        if (title !== urlName) { changes.Title = title; }
        if (hidden) { changes.Hidden = true; }
        return Object.keys(changes).length === 0 ? id
          : this.updateList(id, changes, 'List title or visibility could not be set').then(() => id);
      });
  }

  private ensureEmailColumn(settings: IFeedbackSettings, contactListId: string): Promise<void> {
    const { internalName, displayName } = settings.contactEmailColumn;
    const schemaXml = `<Field Type="Text" Name="${internalName}" StaticName="${internalName}"`
      + ` DisplayName="${escapeXml(displayName)}" MaxLength="254" Required="TRUE" />`;
    return this.ensureField(contactListId, internalName, schemaXml,
      ADD_TO_DEFAULT_CONTENT_TYPE | ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW);
  }

  private ensurePageColumn(settings: IFeedbackSettings, sitePagesListId: string, contactListId: string): Promise<void> {
    const { internalName, displayName } = settings.contactColumn;
    const match = GUID.exec(contactListId);
    if (!match) {
      return Promise.reject(new SetupError('Invalid contact list id', false));
    }
    const schemaXml = `<Field Type="Lookup" Name="${internalName}" StaticName="${internalName}"`
      + ` DisplayName="${escapeXml(displayName)}" List="{${match[1]}}" ShowField="Title" />`;
    return this.ensureField(sitePagesListId, internalName, schemaXml, ADD_TO_ALL_CONTENT_TYPES | ADD_FIELD_INTERNAL_NAME_HINT);
  }

  /** Creates the field unless the list already has a field with this internal name. */
  private ensureField(listId: string, internalName: string, schemaXml: string, options: number): Promise<void> {
    return this.readField(listId, internalName).then((existing: IFieldSchema | undefined) => {
      if (existing) {
        logVerbose('Column exists', { internalName });
        return undefined;
      }
      const body = {
        parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: schemaXml, Options: options }
      };
      logVerbose('Creating column', { schemaXml });
      return this.post(`${listUrl(this.api, listId)}/fields/createfieldasxml`, body, 'Column could not be created')
        .then(() => logInfo('Column created', { internalName }));
    });
  }

  /**
   * $filter on the field collection: a missing field is an empty result, not an error status.
   * listMayBeMissing = true: a missing list (404) also counts as "field missing" (check, not setup).
   */
  private readField(listId: string, internalName: string, listMayBeMissing?: boolean): Promise<IFieldSchema | undefined> {
    const filter = encodeURIComponent(`InternalName eq '${internalName}'`);
    return this.get(`${listUrl(this.api, listId)}/fields?$select=InternalName,TypeAsString,LookupList&$filter=${filter}`, 'Columns could not be read', listMayBeMissing)
      .then((json: unknown) => {
        const field = odataItems(json).filter((f: IUntyped) => f.InternalName === internalName)[0];
        return field ? { type: String(field.TypeAsString || ''), lookupList: String(field.LookupList || '') } : undefined;
      });
  }

  /** Deletes a field of a list (the broken lookup column in Site Pages). */
  private deleteField(listId: string, internalName: string): Promise<unknown> {
    const url = `${listUrl(this.api, listId)}/fields/getbyinternalnameortitle('${internalName}')`;
    return this.post(url, {}, 'Column could not be deleted', { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' })
      .then(() => logInfo('Broken lookup column deleted', { internalName }));
  }

  /** notFoundIsEmpty = true: 404 resolves with undefined instead of an error (and is not logged as error). */
  private get(url: string, message: string, notFoundIsEmpty?: boolean): Promise<unknown> {
    logVerbose('GET ' + url);
    return this.httpClient.get(url, SPHttpClient.configurations.v1, NOMETADATA)
      .then((response: SPHttpClientResponse) => {
        if (notFoundIsEmpty && response.status === 404) {
          logVerbose('Not found: ' + url);
          return undefined;
        }
        return this.checked(response, message);
      });
  }

  private post(url: string, body: unknown, message: string, extraHeaders?: { [name: string]: string }): Promise<unknown> {
    logVerbose('POST ' + url);
    return this.httpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { ...VERBOSE_HEADERS, ...(extraHeaders || {}) },
      body: JSON.stringify(body)
    }).then((response: SPHttpClientResponse) => this.checked(response, message));
  }

  private checked(response: SPHttpClientResponse, message: string): Promise<unknown> {
    if (response.ok) {
      // MERGE answers with 204 (no body).
      return response.status === 204 ? Promise.resolve(undefined) : response.json().catch(() => undefined);
    }
    return response.text().then((text: string) => {
      logError(message + ', status ' + response.status, { response: text });
      throw new SetupError(message, response.status === 401 || response.status === 403);
    });
  }
}

/** A list of a REST answer (with RootFolder expanded); undefined without a valid ID. */
function toListInfo(list: IUntyped): IListInfo | undefined {
  const id = String(list.Id || '');
  return GUID.test(id) ? {
    id,
    title: typeof list.Title === 'string' ? list.Title.trim() : '',
    serverRelativeUrl: String(odataEntity(list.RootFolder).ServerRelativeUrl || ''),
    hidden: list.Hidden === true
  } : undefined;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
