import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackSettings, IPageInfo, IRecipientDefaults, IRecipientSource, ISiteSettings, RecipientStatus } from '../models/IFeedback';
import { logVerbose, logWarning } from '../utils/logging';
import { validEmails } from '../utils/sanitize';
import { asArray, IUntyped, odataEntity, odataItems } from '../utils/untyped';

export const GUID = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i;

/** JSON without metadata (OData v3 format), supported by SharePoint 2016/2019 and Online. */
export const NOMETADATA = { headers: { 'Accept': 'application/json;odata=nometadata', 'OData-Version': '' } };

const LOOKUP_TYPES = ['Lookup', 'LookupMulti'];

/** Schema of a field, undefined if the list has no field with this internal name. */
interface IFieldSchema {
  type: string;
  lookupList: string;
}

/** Values read from SharePoint before validation, or a status that ends the evaluation early. */
type RawResult = string[] | RecipientStatus;

/**
 * Reads the default recipients of the form:
 *
 * - Contact person: lookup column of the current page (Site Pages library, config contact.pageColumn or
 *   settings list). The e-mail column (config contact.emailColumn) of the selected entry(s) in the contact list is read.
 * - Shared mailbox (config mailbox): entry of the settings list (already read by SettingsService) or value
 *   of the property bag of the current web.
 *
 * The result carries a status, so that the form can tell "nothing stored for this page" from
 * "column missing" and offer the setup. Read only when the dialog opens, then cached for the page view.
 */
export class RecipientService {
  private readonly api: string;
  private contactCache: { [pageKey: string]: Promise<IRecipientSource> } = {};
  private propertyBagCache: Promise<IRecipientSource> | undefined;

  constructor(
    private readonly httpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly policy: ISecurityPolicy
  ) {
    this.api = webAbsoluteUrl.replace(/\/+$/, '') + '/_api';
  }

  public load(page: IPageInfo, site: ISiteSettings): Promise<IRecipientDefaults> {
    return Promise.all([this.getContact(page, site.settings), this.getMailbox(site)])
      .then(([contact, mailbox]: [IRecipientSource, IRecipientSource]) => ({ contact, mailbox }));
  }

  /** Forgets cached results, e.g. after a column was created or the mailbox was stored. */
  public reset(): void {
    this.contactCache = {};
    this.propertyBagCache = undefined;
  }

  private getContact(page: IPageInfo, settings: IFeedbackSettings): Promise<IRecipientSource> {
    const column = settings.contactColumn.internalName;
    if (!column) {
      return Promise.resolve(source('notConfigured'));
    }
    // The column is part of the key: it can be changed in the settings list.
    const key = page.listId + '|' + page.itemId + '|' + column + '|' + settings.contactEmailColumn.internalName;
    if (!this.contactCache[key]) {
      this.contactCache[key] = this.readContact(page, settings)
        .catch((error: unknown) => {
          logWarning('Contact person could not be read', { column, error: String(error) });
          // Do not keep the failure: the next opening tries again.
          delete this.contactCache[key];
          return source('error');
        });
    }
    return this.contactCache[key];
  }

  /**
   * 1. Schema of the lookup column in Site Pages (missing? lookup? which list?)
   * 2. In parallel: IDs of the selected entries and the e-mail column of the contact list
   * 3. The e-mail addresses of these entries
   */
  private readContact(page: IPageInfo, settings: IFeedbackSettings): Promise<IRecipientSource> {
    const column = settings.contactColumn.internalName;
    const emailColumn = settings.contactEmailColumn.internalName;
    const list = listUrl(this.api, page.listId);
    return this.readFieldSchema(list, column).then((schema: IFieldSchema | undefined): IRecipientSource | Promise<IRecipientSource> => {
      if (!schema) {
        return this.evaluate('missing', 'contact');
      }
      const match = GUID.exec(schema.lookupList);
      if (LOOKUP_TYPES.indexOf(schema.type) < 0 || !match) {
        return this.evaluate('wrongType', 'contact');
      }
      const lookupListId = match[1];
      const contactList = listUrl(this.api, lookupListId);
      return Promise.all([
        this.getJson(`${list}/items(${page.itemId})?$select=${column}Id`),
        this.readFieldSchema(contactList, emailColumn)
      ]).then(([item, emailSchema]: [unknown, IFieldSchema | undefined]): RawResult | Promise<RawResult> => {
        if (!emailSchema) {
          return 'lookupEmailMissing';
        }
        const ids = asArray(odataEntity(item)[column + 'Id'])
          .filter((id: unknown): id is number => typeof id === 'number' && id > 0 && Math.floor(id) === id)
          .slice(0, this.policy.maxRecipientsPerField);
        logVerbose('Contact entries', { lookupList: lookupListId, ids });
        if (ids.length === 0) {
          return [];
        }
        const filter = encodeURIComponent(ids.map((id: number) => `Id eq ${id}`).join(' or '));
        return this.getJson(`${contactList}/items?$select=${emailColumn}&$filter=${filter}&$top=${ids.length}`)
          .then((json: unknown) => odataItems(json).reduce((all: string[], entry: IUntyped) => all.concat(textOf(entry[emailColumn])), []));
      }).then((raw: RawResult) => ({ ...this.evaluate(raw, 'contact'), lookupListId }));
    });
  }

  private getMailbox(site: ISiteSettings): Promise<IRecipientSource> {
    const { key, storage } = site.settings.mailbox;
    if (!key) {
      return Promise.resolve(source('notConfigured'));
    }
    if (storage === 'list') {
      // stored undefined: the settings list could not be read (logged by SettingsService).
      return Promise.resolve(site.stored ? this.evaluateMailbox(textOf(site.stored[key])) : source('error'));
    }
    if (!this.propertyBagCache) {
      this.propertyBagCache = this.readMailboxFromPropertyBag(key)
        .then((values: string[]) => this.evaluateMailbox(values))
        .catch((error: unknown) => {
          logWarning('Shared mailbox could not be read', { key, error: String(error) });
          this.propertyBagCache = undefined;
          return source('error');
        });
    }
    return this.propertyBagCache;
  }

  private evaluateMailbox(values: string[]): IRecipientSource {
    const nonEmpty = values.filter((value: string) => !!value.trim());
    return this.evaluate(nonEmpty.length > 0 ? nonEmpty : 'missing', 'mailbox');
  }

  /** Property bag: $select reads only this one value. A missing key returns an empty object. */
  private readMailboxFromPropertyBag(key: string): Promise<string[]> {
    const name = encodePropertyName(key);
    return this.getJson(`${this.api}/web/AllProperties?$select=${name}`)
      .then((json: unknown) => textOf(odataEntity(json)[name]));
  }

  /** Turns raw values into a status: no values -> empty, values without a valid address -> invalid. */
  private evaluate(raw: RawResult, origin: string): IRecipientSource {
    if (typeof raw === 'string') {
      logVerbose('Recipient source: ' + raw, { origin });
      return source(raw);
    }
    const emails = validEmails(raw, origin, this.policy);
    const nonEmpty = raw.filter((value: string) => !!value.trim());
    return { status: emails.length > 0 ? 'ok' : (nonEmpty.length > 0 ? 'invalid' : 'empty'), emails };
  }

  /** $filter on the field collection: a missing field is an empty result, not an error status. */
  private readFieldSchema(list: string, internalName: string): Promise<IFieldSchema | undefined> {
    const filter = encodeURIComponent(`InternalName eq '${internalName}'`);
    return this.getJson(`${list}/fields?$select=InternalName,TypeAsString,LookupList&$filter=${filter}`)
      .then((json: unknown) => {
        const field = odataItems(json).filter((f: IUntyped) => f.InternalName === internalName)[0];
        return field ? { type: String(field.TypeAsString || ''), lookupList: String(field.LookupList || '') } : undefined;
      });
  }

  private getJson(url: string): Promise<unknown> {
    logVerbose('GET ' + url);
    return this.httpClient.get(url, SPHttpClient.configurations.v1, NOMETADATA)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          const error = new Error('Status ' + response.status + ' for ' + url) as Error & { status?: number };
          error.status = response.status;
          throw error;
        }
        return response.json();
      });
  }
}

function source(status: RecipientStatus): IRecipientSource {
  return { status, emails: [] };
}

/** REST URL of a list by its ID. Throws for an invalid ID (it is inserted into the URL). */
export function listUrl(api: string, listId: string): string {
  const match = GUID.exec(listId);
  if (!match) {
    throw new Error('Invalid list id');
  }
  return `${api}/web/lists(guid'${match[1]}')`;
}

/** REST URL of a list by its URL relative to the web ("Lists/<name>"), SharePoint 2016 and later. */
export function settingsListUrl(api: string, webServerRelativeUrl: string, relativeListUrl: string): string {
  const serverRelative = webServerRelativeUrl.replace(/\/+$/, '') + '/' + relativeListUrl;
  return `${api}/web/getlist('${encodeURIComponent(serverRelative.replace(/'/g, "''"))}')`;
}

/** Text value; may contain several addresses separated by ";" (split by validEmails). */
function textOf(value: unknown): string[] {
  return typeof value === 'string' ? [value] : [];
}

/**
 * REST returns property bag keys as OData property names: every character other than a letter or digit
 * is encoded as _xHHHH_ ("My_Key" -> "My_x005f_Key", "My Key" -> "My_x0020_Key").
 */
function encodePropertyName(key: string): string {
  return key.replace(/[^A-Za-z0-9]/g, (c: string) => '_x' + ('000' + c.charCodeAt(0).toString(16)).slice(-4) + '_');
}
