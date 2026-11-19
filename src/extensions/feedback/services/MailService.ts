import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { logError, logInfo, logVerbose } from '../utils/logging';
import { IMail } from '../utils/mail';

export class MailService {
  private readonly endpoint: string;

  constructor(private readonly httpClient: SPHttpClient, webAbsoluteUrl: string) {
    this.endpoint = webAbsoluteUrl.replace(/\/+$/, '') + '/_api/SP.Utilities.Utility.SendEmail';
  }

  public send(mail: IMail): Promise<void> {
    const properties: { [key: string]: unknown } = {
      __metadata: { type: 'SP.Utilities.EmailProperties' },
      To: { results: mail.to },
      Subject: mail.subject,
      Body: mail.body
    };
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
