import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { logError, logInfo, logVerbose } from '../utils/logging';
import { IMail } from '../utils/mail';

/**
 * Sends mails via the SharePoint REST endpoint SP.Utilities.Utility.SendEmail.
 *
 * - The sender is SharePoint (the outgoing mail address of the farm or tenant).
 * - SharePoint 2019: outgoing e-mail (SMTP) must be configured in Central Administration.
 * - SharePoint Online: only recipients inside the organization receive the mail.
 */
export class MailService {
  private readonly endpoint: string;

  constructor(private readonly httpClient: SPHttpClient, webAbsoluteUrl: string) {
    this.endpoint = webAbsoluteUrl.replace(/\/+$/, '') + '/_api/SP.Utilities.Utility.SendEmail';
  }

  /** Resolves when SharePoint has accepted the mail, rejects otherwise (details are logged, not shown). */
  public send(mail: IMail): Promise<void> {
    // The endpoint expects the verbose format with __metadata. SPHttpClient sets "OData-Version: 4.0"
    // by default, under which __metadata is rejected - the empty header suppresses it.
    const properties: { [key: string]: unknown } = {
      __metadata: { type: 'SP.Utilities.EmailProperties' },
      To: { results: mail.to },
      Subject: mail.subject,
      Body: mail.body
    };
    // Optional parts only if set: empty collections or an empty sender can be rejected.
    if (mail.cc.length > 0) {
      properties.CC = { results: mail.cc };
    }
    if (mail.bcc.length > 0) {
      properties.BCC = { results: mail.bcc };
    }
    if (mail.from) {
      properties.From = mail.from;
    }
    if (mail.replyTo) {
      properties.AdditionalHeaders = {
        __metadata: { type: 'Collection(SP.KeyValue)' },
        results: [{ __metadata: { type: 'SP.KeyValue' }, Key: 'Reply-To', Value: mail.replyTo, ValueType: 'Edm.String' }]
      };
    }
    const body = { properties };
    logVerbose('Sending mail', { to: mail.to, cc: mail.cc, bcc: mail.bcc, from: mail.from, replyTo: mail.replyTo, subject: mail.subject });

    return this.httpClient.post(this.endpoint, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'OData-Version': ''
      },
      body: JSON.stringify(body)
    }).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        logInfo('Mail sent', { recipients: mail.to.length });
        return;
      }
      return response.text().then((text: string) => {
        logError('Mail could not be sent, status ' + response.status, { response: text });
        throw new Error('SendEmail failed with status ' + response.status);
      });
    }, (error: unknown) => {
      logError('Mail could not be sent', { error: String(error) });
      throw error;
    });
  }
}
