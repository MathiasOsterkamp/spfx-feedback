import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { NOMETADATA, settingsListUrl } from './RecipientService';
import { ISecurityPolicy } from '../config/securityConfig';
import { IFeedbackSettings, ISiteSettings, IStoredValues } from '../models/IFeedback';
import { logVerbose, logWarning } from '../utils/logging';
import { applyStoredValues } from '../utils/sanitize';
import { IUntyped, odataItems } from '../utils/untyped';

/** More entries than the settings list ever needs; protects against a list that was filled by mistake. */
const MAX_ENTRIES = 500;

/**
 * Reads the settings list of the site (config settingsList) and applies its entries to the code
 * configuration: contact list and column, categories, mail template, shared mailbox (storage "list").
 *
 * One request when the dialog opens, cached for the page view. Never fails: if the list cannot be read,
 * the code configuration applies and stored is undefined. A missing list or value column means
 * "nothing stored" - the setup view creates both when it saves.
 */
export class SettingsService {
  private readonly api: string;
  private cache: Promise<ISiteSettings> | undefined;

  constructor(
    private readonly httpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly webServerRelativeUrl: string,
    /** Checked code configuration (config/feedbackConfig.ts), without the values of the settings list. */
    public readonly base: IFeedbackSettings,
    private readonly policy: ISecurityPolicy
  ) {
    this.api = webAbsoluteUrl.replace(/\/+$/, '') + '/_api';
  }

  public load(): Promise<ISiteSettings> {
    if (!this.cache) {
      this.cache = this.readStored().then((stored: IStoredValues | undefined): ISiteSettings => {
        const settings = stored ? applyStoredValues(this.base, stored, this.policy) : this.base;
        logVerbose('Site settings', { stored, settings });
        return { settings, stored };
      });
    }
    return this.cache;
  }

  /** Forgets the cached values, e.g. after the setup view saved. */
  public reset(): void {
    this.cache = undefined;
  }

  /** Only the known keys are kept; for duplicate titles the first entry (lowest ID) counts. */
  private readStored(): Promise<IStoredValues | undefined> {
    const { url, valueColumn, keys } = this.base.settingsList;
    if (!url) {
      // Invalid configuration, already logged.
      return Promise.resolve({});
    }
    const column = valueColumn.internalName;
    const known = [keys.contactList, keys.contactColumn, keys.categories, keys.mailSubject, keys.mailBody, this.base.mailbox.key]
      .filter((key: string) => !!key);
    const request = `${settingsListUrl(this.api, this.webServerRelativeUrl, url)}/items?$select=Title,${column}&$orderby=Id&$top=${MAX_ENTRIES}`;
    logVerbose('GET ' + request);
    return this.httpClient.get(request, SPHttpClient.configurations.v1, NOMETADATA)
      .then((response: SPHttpClientResponse): IStoredValues | Promise<IStoredValues> => {
        if (response.status === 404 || response.status === 400) {
          logVerbose('Settings list or value column missing', { status: response.status });
          return {};
        }
        if (!response.ok) {
          throw new Error('Status ' + response.status + ' for ' + request);
        }
        return response.json().then((json: unknown) => {
          const stored: IStoredValues = {};
          odataItems(json).forEach((entry: IUntyped) => {
            const key = typeof entry.Title === 'string' ? entry.Title.trim() : '';
            if (known.indexOf(key) >= 0 && !Object.prototype.hasOwnProperty.call(stored, key)) {
              stored[key] = typeof entry[column] === 'string' ? entry[column] as string : '';
            }
          });
          return stored;
        });
      })
      .catch((error: unknown) => {
        logWarning('Settings list could not be read, the code configuration applies', { list: url, error: String(error) });
        return undefined;
      });
  }
}
