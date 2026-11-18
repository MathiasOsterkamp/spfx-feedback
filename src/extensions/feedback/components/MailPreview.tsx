import * as React from 'react';
import { usePreviewClassNames } from './styles';
import { IMail } from '../utils/mail';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface IMailPreviewProps {
  /** The mail exactly as it will be sent (buildMail). */
  mail: IMail;
}

/**
 * Preview of the notification mail before it is sent: header like in a mail program (subject, recipients,
 * copy, reply-to) and the body as the recipients see it. Not shown: the sender (fixed by SharePoint) and BCC
 * (those recipients are hidden from everyone, also from the reporter).
 */
export const MailPreview: React.FunctionComponent<IMailPreviewProps> = (props: IMailPreviewProps) => {
  const { mail } = props;
  const classNames = usePreviewClassNames();

  // Links in the body would replace the SharePoint page (and the form): open them in a new window instead.
  const onBodyClick = (event: React.MouseEvent<HTMLDivElement>): void => {
    let element = event.target as HTMLElement | null;
    while (element && element !== event.currentTarget && element.tagName !== 'A') {
      element = element.parentElement;
    }
    const href = element && element.tagName === 'A' ? element.getAttribute('href') : null;
    if (href) {
      event.preventDefault();
      if (/^https?:/i.test(href)) {
        window.open(href, '_blank', 'noopener,noreferrer');
      } else if (/^mailto:/i.test(href)) {
        window.location.href = href;
      }
    }
  };

  /** Addresses as chips; an empty list shows "none". */
  const addresses = (emails: string[]): React.ReactNode => (emails.length === 0
    ? <span className={classNames.none}>{strings.previewNone}</span>
    : emails.map((email: string) => <span key={email} className={classNames.address}>{email}</span>));

  const row = (label: string, value: React.ReactNode, valueClassName?: string): React.ReactNode => (
    <tr>
      <th scope="row" className={classNames.label}>{label}</th>
      <td className={valueClassName || classNames.value}>{value}</td>
    </tr>
  );

  return (
    <section aria-labelledby="spfx-feedback-preview-title">
      <h3 id="spfx-feedback-preview-title" className={classNames.title}>{strings.previewHeading}</h3>
      <p className={classNames.hint}>{strings.previewHint}</p>
      <div className={classNames.mail}>
        <table className={classNames.header}>
          <tbody>
            {row(strings.previewSubject, mail.subject, classNames.subject)}
            {row(strings.previewTo, addresses(mail.to))}
            {row(strings.previewCc, addresses(mail.cc))}
            {mail.replyTo && row(strings.previewReplyTo, addresses([mail.replyTo]))}
          </tbody>
        </table>
        {/* The body is built by buildMail from the plain-text template: template text and all values are
            HTML-escaped, the only tags are the ones added there (frame, <br>, links to http(s) and mailto). */}
        {/* eslint-disable-next-line jsx-a11y/click-events-have-key-events, jsx-a11y/no-static-element-interactions -- only intercepts clicks on the links inside, which are keyboard accessible themselves */}
        <div className={classNames.body} onClick={onBodyClick} dangerouslySetInnerHTML={{ __html: mail.body }} />
      </div>
    </section>
  );
};
