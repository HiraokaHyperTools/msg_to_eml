import MsgReader, { FieldsData } from '@kenjiuno/msgreader'
import { decompressRTF } from '@kenjiuno/decompressrtf';
import iconvLite from 'iconv-lite';
import { deEncapsulateSync } from 'rtf-stream-parser';
import MailComposer from 'nodemailer/lib/mail-composer';
import { Buffer } from 'buffer';
import path from 'path';

function changeFileExtension(fileName: string, newExt: string): string {
  const parsed = path.parse(fileName);
  return (parsed.dir ? parsed.dir + path.sep : "") + parsed.name + newExt;
}

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

function convertVLines(vLines: [string | string[], string | string[]][]): string {
  function toText(vCell: string[] | string): string {
    return Array.isArray(vCell)
      ? vCell.join(';')
      : vCell + ""
  }

  function vEscape(text: string): { more: string, text: string } {
    if (text.match(/\r|\n|\t/)) {
      return {
        more: ";ENCODING=QUOTED-PRINTABLE",
        text: text
          .replace(/\t/g, "=09")
          .replace(/\r/g, "=0D")
          .replace(/\n/g, "=0A")
        ,
      };
    }
    else {
      return { more: "", text };
    }
  }

  const lines = [];
  vLines.forEach(
    vLine => {
      if (vLine[1] === undefined) {
        return;
      }
      const printer = vEscape(toText(vLine[1]));

      lines.push(toText(vLine[0]) + printer.more + ":" + printer.text);
    }
  )
  return lines.join("\n");
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

export class ParsedMsg {
  private _msgReader: MsgReader;
  private _msgInfo: FieldsData;
  private _options: MsgConverterOptions;

  constructor(msgReader: MsgReader, msgInfo: FieldsData, options: MsgConverterOptions) {
    this._msgReader = msgReader;
    this._msgInfo = msgInfo;
    this._options = options;
  }

  public get msgReader() { return this._msgReader; }
  public get msgInfo() { return this._msgInfo; }

  /**
   * Check `.msg` file usage
   * 
   * @returns
   * - This will return `IPM.Contact` for contact.
   * - This will return `IPM.Note` for EML.
   */
  public get messageClass(): string { return this._msgInfo.messageClass; }

  /**
   * Assume this is a contact and then convert to vCard.
   * 
   * @returns vCard text
   */
  async toVCardStr(): Promise<string> {
    const source = this._msgInfo;

    const makers = [
      {
        kind: "N",
        provider: () => [
          source.surname,
          source.givenName,
          source.middleName,
          source.displayNamePrefix,
          source.generation,
        ],
      }, {
        kind: "FN",
        provider: () => source.name,
      }, {
        kind: "X-MS-N-YOMI",
        provider: () => [source.yomiLastName, source.yomiFirstName],
      }, {
        kind: "ORG",
        provider: () => [source.companyName, source.departmentName],
      }, {
        kind: "X-MS-ORG-YOMI",
        provider: () => source.yomiCompanyName,
      }, {
        kind: "TITLE",
        provider: () => source.title,
      }, {
        kind: ["TEL", "WORK", "VOICE"],
        provider: () => source.businessTelephoneNumber,
      }, {
        kind: ["TEL", "HOME", "VOICE"],
        provider: () => source.homeTelephoneNumber,
      }, {
        kind: ["TEL", "CELL", "VOICE"],
        provider: () => source.mobileTelephoneNumber,
      }, {
        kind: ["TEL", "WORK", "FAX"],
        provider: () => source.businessFaxNumber,
      }, {
        kind: ["ADR", "WORK", "PREF"],
        provider: () => [
          source.workAddressStreet,
          source.workAddressCity,
          source.workAddressState,
          source.workAddressPostalCode,
          source.workAddressCountry
        ]
      }, {
        kind: ["LABEL", "WORK", "PREF"],
        provider: () => source.workAddress,
      }, {
        kind: ["URL", "WORK"],
        provider: () => source.businessHomePage,
      }, {
        kind: ["EMAIL", "PREF", "INTERNET"],
        provider: () => source.email1EmailAddress,
      }
    ];

    const vLines: [string | string[], string | string[]][] = [];
    vLines.push([["BEGIN"], "VCARD"]);
    vLines.push([["VERSION"], "2.1"]);

    makers.forEach(
      maker => {
        vLines.push([
          maker.kind,
          maker.provider()
        ]);
      }
    )

    vLines.push([["END"], "VCARD"]);

    return convertVLines(vLines);
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML text
   */
  async toEmlStr(): Promise<string> {
    return (await this.toEmlBuffer()).toString('utf-8');
  }

  /**
   * Assume this is a mail message and then convert to EML.
   * 
   * @returns EML file
   */
  async toEmlBuffer(): Promise<Buffer> {
    return (await this.toEmlFrom(this._msgInfo));
  }

  private async toEmlFrom(msgInfo: FieldsData): Promise<Buffer> {
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

    const entity = {
      baseBoundary: this._options.baseBoundary,

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
        "Message-ID": this._options.messageId,
      }
    };

    for (let attachment of attachments) {
      if (attachment.innerMsgContent) {
        const emlBuf = await this.toEmlFrom(attachment.innerMsgContentFields);

        attachmentsRefined.push({
          filename: changeFileExtension(attachment.name ?? "unnamed", ".eml"),
          content: emlBuf,
          cid: attachment.pidContentId,
          contentTransferEncoding: '8bit',
        })
      }
      else {
        const exported = this._msgReader.getAttachment(attachment);
        attachmentsRefined.push({
          filename: exported.fileName,
          content: Buffer.from(exported.content),
          cid: attachment.pidContentId,
        });
      }
    }

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

export class MsgConverter {
  options: MsgConverterOptions;

  constructor(options: MsgConverterOptions) {
    this.options = options || {};
  }

  /**
   * Parse at first
   * 
   * @param input msg file
   * @returns parsed entity
   */
  async parse(input: Uint8Array): Promise<ParsedMsg> {
    const inputMsg = new MsgReader(input);
    inputMsg.parserConfig = {
      ansiEncoding: this.options.ansiEncoding,
    }
    const msgInfo = inputMsg.getFileData();
    if (msgInfo.error) {
      throw new Error(`MSG file parser error: ${msgInfo.error}`);
    }

    return new ParsedMsg(inputMsg, msgInfo, this.options);
  }

  /**
   * Get EML
   * 
   * @param input msg file
   * @returns EML text
   */
  async convertToString(input: Uint8Array): Promise<string> {
    return (await this.parse(input)).toEmlStr();
  }

  /**
   * Get EML
   * 
   * @param input msg file
   * @returns EML file
   */
  async convertToBuffer(input: Uint8Array): Promise<Buffer> {
    return (await this.parse(input)).toEmlBuffer();
  }
}
