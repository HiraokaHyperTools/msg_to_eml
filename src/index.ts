import MsgReader from '@kenjiuno/msgreader'
import { decompressRTF } from '@kenjiuno/decompressrtf';
import iconvLite from 'iconv-lite';
import { deEncapsulateSync } from 'rtf-stream-parser';
import MailComposer from 'nodemailer/lib/mail-composer';
import { Buffer } from 'buffer';
import path from 'path';

async function convertRtfToHtml(rtf: string): Promise<string> {
  if (typeof rtf === "string" && rtf.length !== 0) {
    try {
      const result = deEncapsulateSync(
        rtf,
        {
          mode: "html",
          decode: iconvLite.decode,
        }
      );
      return result.text as string;
    }
    catch {
      //throw new Error("Conversion from RTF to HTML failed.\n" + ex);
    }
  }
  return undefined;
}

function uncompressRtf(compressedRtf?: Uint8Array): string {
  if (compressedRtf === undefined) {
    return undefined;
  }
  return new TextDecoder("utf-8").decode(
    Uint8Array.from(
      decompressRTF([...compressedRtf])
    )
  );
}

function formatFrom(senderName: string, senderEmail: string): string {
  if (senderName) {
    return `${senderName} <${senderEmail}>`;
  }
  else {
    return `${senderEmail}`;
  }
}

export interface MsgConverterOptions {
  /**
   * fixed baseBoundary for EML for testing
   */
  baseBoundary?: string;

  /**
   * fixed messageId for EML for testing
   */
  messageId?: string;

  /**
   * ANSI encoding decoded by iconv-lite
   * 
   * like: `latin1`, `cp1251`, `cp932`, and so on
   */
  ansiEncoding?: string;
}

export class MsgConverter {
  options: MsgConverterOptions;

  constructor(options: MsgConverterOptions) {
    this.options = options || {};
  }

  async convertToString(input: Uint8Array): Promise<string> {
    return (await this.convertToBuffer(input)).toString('utf-8');
  }

  async convertToBuffer(input: Uint8Array): Promise<Buffer> {
    const inputMsg = new MsgReader(input);
    inputMsg.parserConfig = {
      ansiEncoding: this.options.ansiEncoding,
    }
    const msgInfo = inputMsg.getFileData();
    if (msgInfo.error) {
      throw new Error(`MSG file parser error: ${msgInfo.error}`);
    }

    const htmlBody: string = await convertRtfToHtml(
      uncompressRtf(msgInfo.compressedRtf)
    );

    const senderName = msgInfo.senderName;
    const senderEmail = msgInfo.senderEmail;
    const recipients = msgInfo.recipients;
    const subject = msgInfo.subject;
    const attachments = msgInfo.attachments;
    const body = msgInfo.body;

    function applyFallbackRecipients(array, fallback) {
      if (array.length === 0) {
        array.push(fallback);
      }
      return array;
    }

    const attachmentsRefined = [];
    for (let attachment of attachments) {
      const exported = inputMsg.getAttachment(attachment);
      if (attachment.innerMsgContent) {
        const emlBuf = await this.convertToBuffer(exported.content);

        attachmentsRefined.push({
          filename: changeExtension(exported.fileName, ".eml"),
          content: emlBuf,
          cid: attachment.pidContentId,
          contentTransferEncoding: '8bit',
        })
      }
      else {
        attachmentsRefined.push({
          filename: exported.fileName,
          content: Buffer.from(exported.content),
          cid: attachment.pidContentId,
        });
      }
    }

    const entity = {
      baseBoundary: this.options.baseBoundary,

      from: formatFrom(senderName, senderEmail),
      to: applyFallbackRecipients(
        recipients
          .map(
            ({ name, email, recipType }) => {
              return recipType === "to" ? { name, email } : null
            }
          )
          .filter((entry) => entry !== null), { name: "undisclosed-recipients" }),
      cc: recipients
        .map(({ name, email, recipType }) =>
          recipType === "cc" ? { name, email } : null
        )
        .filter((entry) => entry !== null),
      bcc: recipients
        .map(({ name, email, recipType }) =>
          recipType === "bcc" ? { name, email } : null
        )
        .filter((entry) => entry !== null),
      subject,
      text: body,
      html: htmlBody,
      attachments: attachmentsRefined,
      headers: {
        "Date": false
          || msgInfo.messageDeliveryTime
          || msgInfo.clientSubmitTime
          || msgInfo.lastModificationTime
          || msgInfo.creationTime,
        "Message-ID": this.options.messageId,
      }
    };

    return await new Promise((resolve, reject) => {
      try {
        const mail = new MailComposer(entity);
        mail.compile().build(function (error, message) {
          if (error) {
            reject(new Error("EML composition failed.\n" + error + "\n\n" + error.stack));
            return;
          }
          else {
            resolve(message);
          }
        });
      } catch (ex) {
        reject(new Error("EML composition failed.\n" + ex));
      }
    });
  }
}

function changeExtension(fileName: string, newExt: string): string {
  const parsed = path.parse(fileName);
  return (parsed.dir ? parsed.dir + path.sep : "") + parsed.name + newExt;
}
