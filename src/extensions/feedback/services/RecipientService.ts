import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackSettings, IPageInfo, IRecipientDefaults, IRecipientSource, ISiteSettings, RecipientStatus } from '../models/IFeedback';
import { logVerbose, logWarning } from '../utils/logging';
import { validEmails } from '../utils/sanitize';
import { asArray, IUntyped, odataEntity, odataItems } from '../utils/untyped';

export const GUID = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i;

/** OData v3 JSON without metadata (SharePoint 2016/2019 and Online). */
export const NOMETADATA = { headers: { 'Accept': 'application/json;odata=nometadata', 'OData-Version': '' } };

const LOOKUP_TYPES = ['Lookup', 'LookupMulti'];

interface IFieldSchema {
  type: string;
  lookupList: string;
}

type RawResult = string[] | RecipientStatus;

/** The status lets the form tell "nothing stored" from "column missing" and offer the setup. */
export class RecipientService {
  private readonly api: string;
  private contactCache: { [pageKey: string]: Promise<IRecipientSource> } = {};

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

  public reset(): void {
    this.contactCache = {};
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
          // 404: the lookup target list was deleted - fixed by the setup.
          if ((error as { status?: number }).status === 404) {
            logWarning('The list of the contact lookup column does not exist, run the setup (gear icon)', { column });
            return source('missing');
          }
          logWarning('Contact person could not be read', { column, error: String(error) });
          // Not cached: the next opening tries again.
          delete this.contactCache[key];
          return source('error');
        });
    }
    return this.contactCache[key];
  }

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
      }).then((raw: RawResult) => this.evaluate(raw, 'contact'));
    });
  }

  private getMailbox(site: ISiteSettings): Promise<IRecipientSource> {
    const key = site.settings.mailbox.key;
    if (!key) {
      return Promise.resolve(source('notConfigured'));
    }
    if (!site.stored) {
      return Promise.resolve(source('error'));
    }
    const nonEmpty = textOf(site.stored[key]).filter((value: string) => !!value.trim());
    return Promise.resolve(this.evaluate(nonEmpty.length > 0 ? nonEmpty : 'missing', 'mailbox'));
  }

  private evaluate(raw: RawResult, origin: string): IRecipientSource {
    if (typeof raw === 'string') {
      logVerbose('Recipient source: ' + raw, { origin });
      return source(raw);
    }
    const emails = validEmails(raw, origin, this.policy);
    const nonEmpty = raw.filter((value: string) => !!value.trim());
    return { status: emails.length > 0 ? 'ok' : (nonEmpty.length > 0 ? 'invalid' : 'empty'), emails };
  }

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

/** Throws for an invalid ID (it is inserted into the URL). */
export function listUrl(api: string, listId: string): string {
  const match = GUID.exec(listId);
  if (!match) {
    throw new Error('Invalid list id');
  }
  return `${api}/web/lists(guid'${match[1]}')`;
}

/** getlist() needs SharePoint 2016 or later. */
export function settingsListUrl(api: string, webServerRelativeUrl: string, relativeListUrl: string): string {
  const serverRelative = webServerRelativeUrl.replace(/\/+$/, '') + '/' + relativeListUrl;
  return `${api}/web/getlist('${encodeURIComponent(serverRelative.replace(/'/g, "''"))}')`;
}

function textOf(value: unknown): string[] {
  return typeof value === 'string' ? [value] : [];
}
