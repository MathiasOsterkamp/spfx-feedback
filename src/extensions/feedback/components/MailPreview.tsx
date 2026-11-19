import * as React from 'react';
import { usePreviewClassNames } from './styles';
import { IMail } from '../utils/mail';
import * as strings from 'FeedbackApplicationCustomizerStrings';

export interface IMailPreviewProps {
  mail: IMail;
}

export const MailPreview: React.FunctionComponent<IMailPreviewProps> = (props: IMailPreviewProps) => {
  const { mail } = props;
  const classNames = usePreviewClassNames();

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
        {/* Safe: buildMail HTML-escapes template and values; its only tags are the frame, <br> and http(s)/mailto links. */}
        {/* eslint-disable-next-line jsx-a11y/click-events-have-key-events, jsx-a11y/no-static-element-interactions -- only intercepts clicks on the links inside, which are keyboard accessible themselves */}
        <div className={classNames.body} onClick={onBodyClick} dangerouslySetInnerHTML={{ __html: mail.body }} />
      </div>
    </section>
  );
};
