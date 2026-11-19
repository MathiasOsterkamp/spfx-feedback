import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { GUID, listUrl, NOMETADATA } from './RecipientService';
import { IFeedbackSettings, IListRef, ISetupStatus, IStoredValues, SetupPartStatus } from '../models/IFeedback';
import { logError, logInfo, logVerbose } from '../utils/logging';
import { defaultSettingEntries, FIELD_NAME, LIST_NAME } from '../utils/sanitize';
import { IUntyped, odataEntity, odataItems } from '../utils/untyped';

export interface IContactListOption {
  name: string;
  title: string;
}

export interface ILookupColumnOption {
  internalName: string;
  title: string;
}

const FIELD_TYPE_LOOKUP = 7;

interface IListInfo extends IListRef {
  hidden: boolean;
}

interface IEntries {
  type: string;
  ids: { [key: string]: number };
}

interface IFieldSchema {
  type: string;
  lookupList: string;
}

const GENERIC_LIST_TEMPLATE = 100;

const LOOKUP_TYPES = ['Lookup', 'LookupMulti'];

const ADD_TO_DEFAULT_CONTENT_TYPE = 1;
const ADD_TO_ALL_CONTENT_TYPES = 4;
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;

const VERBOSE_HEADERS = { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose', 'OData-Version': '' };

export class SetupError extends Error {
  constructor(message: string, public readonly accessDenied: boolean) {
    super(message);
    Object.setPrototypeOf(this, SetupError.prototype);
  }
}

export function isSetupComplete(status: ISetupStatus): boolean {
  return [status.settingsList, status.valueColumn, status.entries, status.contactList, status.emailColumn, status.pageColumn]
    .every((part: SetupPartStatus) => part === 'present');
}

export function canSetUp(status: ISetupStatus): boolean {
  const settingsMissing = [status.settingsList, status.valueColumn, status.entries]
    .some((part: SetupPartStatus) => part === 'missing' || part === 'needsUpdate');
  const contactMissing = status.pageColumn !== 'wrongType' && status.pageColumn !== 'notConfigured'
    && [status.contactList, status.emailColumn, status.pageColumn].some((part: SetupPartStatus) => part === 'missing' || part === 'broken');
  return settingsMissing || contactMissing;
}

export class SetupService {
  private readonly api: string;

  constructor(
    private readonly httpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly webServerRelativeUrl: string,
    private readonly base: IFeedbackSettings
  ) {
    this.api = webAbsoluteUrl.replace(/\/+$/, '') + '/_api';
  }

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
      return (lookupMatch ? this.findListById(lists, lookupMatch[1]) : Promise.resolve(undefined))
        .then((lookupList: IListInfo | undefined) => this.checkParts(settings, pageField, isLookup && !!lookupMatch,
          { settings: settingsList, contact: contactList, lookup: lookupList }));
    });
  }

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
    const broken = isLookup && !lookupList;
    const emailList = lookupList || contactList;
    const expected = Object.keys(defaultSettingEntries(this.base));
    return Promise.all([
      settingsList && valueColumn ? this.readField(settingsList.id, valueColumn, true) : Promise.resolve(undefined),
      emailList && emailColumn ? this.readField(emailList.id, emailColumn, true) : Promise.resolve(undefined),
      settingsList ? this.readEntries(settingsList.id) : Promise.resolve(undefined)
    ]).then(([valueField, emailField, entries]: [IFieldSchema | undefined, IFieldSchema | undefined, IEntries | undefined]): ISetupStatus => {
      const entriesComplete = !!entries && expected.every((key: string) => Object.prototype.hasOwnProperty.call(entries.ids, key));
      const status: ISetupStatus = {
        settingsList: !settingsUrl ? 'notConfigured' : (!settingsList ? 'missing' : (settingsList.hidden ? 'present' : 'needsUpdate')),
        valueColumn: !settingsUrl ? 'notConfigured' : (valueField ? 'present' : 'missing'),
        entries: !settingsUrl ? 'notConfigured' : (entriesComplete ? 'present' : 'missing'),
        contactList: lookupList ? 'present' : (!settings.contactList.url ? 'notConfigured' : (contactList ? 'present' : 'missing')),
        emailColumn: !emailColumn ? 'notConfigured' : (emailField ? 'present' : 'missing'),
        pageColumn: !column ? 'notConfigured' : (!pageField ? 'missing' : (lookupList ? 'present' : (broken ? 'broken' : 'wrongType'))),
        lookupList
      };
      logVerbose('Setup status', status);
      return status;
    });
  }

  private findListById(lists: IListInfo[], id: string): Promise<IListInfo | undefined> {
    const known = lists.filter((list: IListInfo) => list.id.toLowerCase() === id.toLowerCase())[0];
    if (known) {
      return Promise.resolve(known);
    }
    const url = `${listUrl(this.api, id)}?$select=Id,Title,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    return this.get(url, 'List could not be read', true)
      .then((json: unknown) => (json === undefined ? undefined : toListInfo(odataEntity(json))));
  }

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
        ? removeBroken().then(() => this.setupContact(settings, sitePagesListId, status.pageColumn === 'present' ? status.lookupList : undefined))
        : undefined))
      .then(() => logInfo('Setup completed'));
  }

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

  private ensureSettingsList(): Promise<string> {
    const { url, title, valueColumn } = this.base.settingsList;
    const column = valueColumn.internalName;
    const schemaXml = `<Field Type="Note" Name="${column}" StaticName="${column}"`
      + ` DisplayName="${escapeXml(valueColumn.displayName)}" NumLines="6" RichText="FALSE" />`;
    return this.ensureList(url, title, true)
      .then((listId: string) => this.ensureField(listId, column, schemaXml,
        ADD_TO_DEFAULT_CONTENT_TYPE | ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW).then(() => listId));
  }

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

  private setupContact(settings: IFeedbackSettings, sitePagesListId: string, existingLookupList: IListRef | undefined): Promise<void> {
    const listIdPromise = existingLookupList && GUID.test(existingLookupList.id)
      ? Promise.resolve(existingLookupList.id)
      : this.ensureList(settings.contactList.url, settings.contactList.title);
    return listIdPromise
      .then((contactListId: string) => this.ensureEmailColumn(settings, contactListId).then(() => contactListId))
      .then((contactListId: string) => this.ensurePageColumn(settings, sitePagesListId, contactListId));
  }

  private getGenericLists(): Promise<IListInfo[]> {
    const filter = encodeURIComponent(`BaseTemplate eq ${GENERIC_LIST_TEMPLATE}`);
    return this.get(`${this.api}/web/lists?$select=Id,Title,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=${filter}`, 'Lists could not be read')
      .then((json: unknown) => odataItems(json)
        .map(toListInfo)
        .filter((list: IListInfo | undefined): list is IListInfo => !!list));
  }

  private byUrl(lists: IListInfo[], url: string): IListInfo | undefined {
    const serverRelativeUrl = (this.webServerRelativeUrl.replace(/\/+$/, '') + '/' + url).toLowerCase();
    return lists.filter((list: IListInfo) => list.serverRelativeUrl.toLowerCase() === serverRelativeUrl)[0];
  }

  private findList(url: string): Promise<IListInfo | undefined> {
    return this.getGenericLists().then((lists: IListInfo[]) => this.byUrl(lists, url));
  }

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

  private readField(listId: string, internalName: string, listMayBeMissing?: boolean): Promise<IFieldSchema | undefined> {
    const filter = encodeURIComponent(`InternalName eq '${internalName}'`);
    return this.get(`${listUrl(this.api, listId)}/fields?$select=InternalName,TypeAsString,LookupList&$filter=${filter}`, 'Columns could not be read', listMayBeMissing)
      .then((json: unknown) => {
        const field = odataItems(json).filter((f: IUntyped) => f.InternalName === internalName)[0];
        return field ? { type: String(field.TypeAsString || ''), lookupList: String(field.LookupList || '') } : undefined;
      });
  }

  private deleteField(listId: string, internalName: string): Promise<unknown> {
    const url = `${listUrl(this.api, listId)}/fields/getbyinternalnameortitle('${internalName}')`;
    return this.post(url, {}, 'Column could not be deleted', { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' })
      .then(() => logInfo('Broken lookup column deleted', { internalName }));
  }

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
      return response.status === 204 ? Promise.resolve(undefined) : response.json().catch(() => undefined);
    }
    return response.text().then((text: string) => {
      logError(message + ', status ' + response.status, { response: text });
      throw new SetupError(message, response.status === 401 || response.status === 403);
    });
  }
}

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
