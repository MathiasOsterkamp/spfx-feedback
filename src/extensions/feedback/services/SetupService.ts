import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { GUID, listUrl, NOMETADATA } from './RecipientService';
import { IFeedbackSettings, IStoredValues } from '../models/IFeedback';
import { logError, logInfo, logVerbose } from '../utils/logging';
import { IUntyped, odataEntity, odataItems } from '../utils/untyped';

/** Base template of a custom list ("Generic List"). */
const GENERIC_LIST_TEMPLATE = 100;

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

/**
 * Creates what the form needs, for users with the corresponding permissions:
 *  - contact list (config contact.list or settings list) with the columns "Title" and the e-mail column
 *  - lookup column (config contact.pageColumn or settings list) in the Site Pages library, pointing to the contact list
 *  - entries of the settings list (config settingsList; list and value column are created if missing)
 *  - shared mailbox (config mailbox): entry in the settings list or value in the property bag of the web
 *
 * Every step checks first whether its part already exists, so the setup can simply be repeated
 * after a partial failure. SharePoint checks the permissions on the server.
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
   * Sets up the contact person with the current settings of the site (code configuration and settings list):
   *  1. contact list (found by its URL; created if missing)
   *  2. e-mail column in this list - or, if given, in the list the existing lookup column points to
   *  3. lookup column in Site Pages (if missing), added to all content types -> visible in the page properties
   */
  public setupContact(settings: IFeedbackSettings, sitePagesListId: string, existingLookupListId?: string): Promise<void> {
    const listIdPromise = existingLookupListId && GUID.test(existingLookupListId)
      ? Promise.resolve(existingLookupListId)
      : this.ensureList(settings.contactList.url, settings.contactList.title);
    return listIdPromise
      .then((contactListId: string) => this.ensureEmailColumn(settings, contactListId).then(() => contactListId))
      .then((contactListId: string) => this.ensurePageColumn(settings, sitePagesListId, contactListId))
      .then(() => logInfo('Contact setup completed'));
  }

  /** Stores the shared mailbox where the configuration says (settings list or property bag). */
  public setMailbox(value: string): Promise<void> {
    const { storage, key } = this.base.mailbox;
    return storage === 'list' ? this.saveValues({ [key]: value }) : this.setMailboxInPropertyBag(value);
  }

  /**
   * Stores entries in the settings list: creates the list and the value column if missing, then updates
   * the entry whose title is the key (the first one, as SettingsService reads it), or creates it.
   * An empty value clears an existing entry (= default of the code) and creates none.
   * Plain REST: works with "Design" permission and without custom scripts.
   */
  public saveValues(values: IStoredValues): Promise<void> {
    const keys = Object.keys(values).filter((key: string) => !!key);
    const { url, title, valueColumn } = this.base.settingsList;
    if (keys.length === 0) {
      return Promise.resolve();
    }
    if (!url) {
      return Promise.reject(new SetupError('Settings list not configured', false));
    }
    const column = valueColumn.internalName;
    const schemaXml = `<Field Type="Note" Name="${column}" StaticName="${column}"`
      + ` DisplayName="${escapeXml(valueColumn.displayName)}" NumLines="6" RichText="FALSE" />`;
    return this.ensureList(url, title)
      .then((listId: string) => this.ensureField(listId, column, schemaXml,
        ADD_TO_DEFAULT_CONTENT_TYPE | ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW).then(() => listId))
      .then((listId: string) => {
        const list = listUrl(this.api, listId);
        return Promise.all([
          this.get(`${list}?$select=ListItemEntityTypeFullName`, 'List could not be read'),
          this.get(`${list}/items?$select=Id,Title&$orderby=Id&$top=500`, 'Settings could not be read')
        ]).then(([info, items]: [unknown, unknown]) => {
          const type = String(odataEntity(info).ListItemEntityTypeFullName || '');
          const existing: { [key: string]: number } = {};
          odataItems(items).forEach((item: IUntyped) => {
            const key = typeof item.Title === 'string' ? item.Title.trim() : '';
            if (key && typeof item.Id === 'number' && !Object.prototype.hasOwnProperty.call(existing, key)) {
              existing[key] = item.Id;
            }
          });
          // One after the other: only a few entries, and a failure stops at a defined point.
          return keys.reduce((chain: Promise<unknown>, key: string) => chain.then(() => {
            const body: { [name: string]: unknown } = { __metadata: { type }, Title: key };
            body[column] = values[key];
            if (Object.prototype.hasOwnProperty.call(existing, key)) {
              return this.post(`${list}/items(${existing[key]})`, body, 'Setting could not be stored', { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
            }
            return values[key] ? this.post(`${list}/items`, body, 'Setting could not be stored') : undefined;
          }), Promise.resolve<unknown>(undefined));
        });
      })
      .then(() => logInfo('Settings stored', { list: url, keys }));
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

  /** ID of the list with this URL ("Lists/<name>"); the list is created if it does not exist. */
  private ensureList(url: string, title: string): Promise<string> {
    const serverRelativeUrl = (this.webServerRelativeUrl.replace(/\/+$/, '') + '/' + url).toLowerCase();
    const filter = encodeURIComponent(`BaseTemplate eq ${GENERIC_LIST_TEMPLATE}`);
    return this.get(`${this.api}/web/lists?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=${filter}`, 'Lists could not be read')
      .then((json: unknown) => {
        const existing = odataItems(json).filter((list: IUntyped) =>
          String(odataEntity(list.RootFolder).ServerRelativeUrl || '').toLowerCase() === serverRelativeUrl)[0];
        if (existing && GUID.test(String(existing.Id))) {
          logVerbose('List exists', { url });
          return String(existing.Id);
        }
        return this.createList(url.substr('Lists/'.length), title);
      });
  }

  /** The URL of a list comes from the title at creation; a different title is set afterwards. */
  private createList(urlName: string, title: string): Promise<string> {
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
        logInfo('List created', { url: 'Lists/' + urlName });
        if (title === urlName) {
          return id;
        }
        return this.post(listUrl(this.api, id), { __metadata: { type: 'SP.List' }, Title: title }, 'List title could not be set',
          { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }).then(() => id);
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
    const list = listUrl(this.api, listId);
    const filter = encodeURIComponent(`InternalName eq '${internalName}'`);
    return this.get(`${list}/fields?$select=InternalName&$filter=${filter}`, 'Columns could not be read')
      .then((json: unknown) => {
        if (odataItems(json).some((field: IUntyped) => field.InternalName === internalName)) {
          logVerbose('Column exists', { internalName });
          return undefined;
        }
        const body = {
          parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: schemaXml, Options: options }
        };
        logVerbose('Creating column', { schemaXml });
        return this.post(`${list}/fields/createfieldasxml`, body, 'Column could not be created')
          .then(() => logInfo('Column created', { internalName }));
      });
  }

  private get(url: string, message: string): Promise<unknown> {
    logVerbose('GET ' + url);
    return this.httpClient.get(url, SPHttpClient.configurations.v1, NOMETADATA)
      .then((response: SPHttpClientResponse) => this.checked(response, message));
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

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
