"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MsgConverter = exports.ParsedMsg = void 0;
const msgreader_1 = __importDefault(require("@kenjiuno/msgreader"));
const decompressrtf_1 = require("@kenjiuno/decompressrtf");
const iconv_lite_1 = __importDefault(require("iconv-lite"));
const rtf_stream_parser_1 = require("rtf-stream-parser");
const mail_composer_1 = __importDefault(require("nodemailer/lib/mail-composer"));
const buffer_1 = require("buffer");
const path_1 = __importDefault(require("path"));
function changeFileExtension(fileName, newExt) {
    const parsed = path_1.default.parse(fileName);
    return (parsed.dir ? parsed.dir + path_1.default.sep : "") + parsed.name + newExt;
}
async function convertRtfToHtml(rtf) {
    if (typeof rtf === "string" && rtf.length !== 0) {
        try {
            const result = (0, rtf_stream_parser_1.deEncapsulateSync)(rtf, {
                mode: "html",
                decode: iconv_lite_1.default.decode,
            });
            return result.text;
        }
        catch {
            //throw new Error("Conversion from RTF to HTML failed.\n" + ex);
        }
    }
    return undefined;
}
function uncompressRtf(compressedRtf) {
    if (compressedRtf === undefined) {
        return undefined;
    }
    return new TextDecoder("utf-8").decode(Uint8Array.from((0, decompressrtf_1.decompressRTF)([...compressedRtf])));
}
function formatFrom(senderName, senderEmail) {
    if (senderName) {
        return `${senderName} <${senderEmail}>`;
    }
    else {
        return `${senderEmail}`;
    }
}
function convertVLines(vLines) {
    function toText(vCell) {
        return Array.isArray(vCell)
            ? vCell.join(';')
            : vCell + "";
    }
    function vEscape(text) {
        if (text.match(/\r|\n|\t/)) {
            return {
                more: ";ENCODING=QUOTED-PRINTABLE",
                text: text
                    .replace(/\t/g, "=09")
                    .replace(/\r/g, "=0D")
                    .replace(/\n/g, "=0A"),
            };
        }
        else {
            return { more: "", text };
        }
    }
    const lines = [];
    vLines.forEach(vLine => {
        if (vLine[1] === undefined) {
            return;
        }
        const printer = vEscape(toText(vLine[1]));
        lines.push(toText(vLine[0]) + printer.more + ":" + printer.text);
    });
    return lines.join("\n");
}
class ParsedMsg {
    constructor(msgReader, msgInfo, options) {
        this._msgReader = msgReader;
        this._msgInfo = msgInfo;
        this._options = options;
    }
    get msgReader() { return this._msgReader; }
    get msgInfo() { return this._msgInfo; }
    /**
     * Decide convertibility delivered from message class.
     *
     * - vCard, if `IPM.Contact`
     * - EML, if `IPM.Note`
     */
    get messageClass() { return this._msgInfo.messageClass; }
    /**
     * Assume this is a contact and then convert to vCard.
     *
     * @returns vCard text
     */
    async toVCardStr() {
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
        const vLines = [];
        vLines.push([["BEGIN"], "VCARD"]);
        vLines.push([["VERSION"], "2.1"]);
        makers.forEach(maker => {
            vLines.push([
                maker.kind,
                maker.provider()
            ]);
        });
        vLines.push([["END"], "VCARD"]);
        return convertVLines(vLines);
    }
    /**
     * Assume this is a mail message and then convert to EML.
     *
     * @returns EML text
     */
    async toEmlStr() {
        return (await this.toEmlBuffer()).toString('utf-8');
    }
    /**
     * Assume this is a mail message and then convert to EML.
     *
     * @returns EML file
     */
    async toEmlBuffer() {
        return (await this.toEmlFrom(this._msgInfo));
    }
    async toEmlFrom(msgInfo) {
        const htmlBody = await convertRtfToHtml(uncompressRtf(msgInfo.compressedRtf));
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
            to: applyFallbackRecipients(recipients
                .map(({ name, email, recipType }) => {
                return recipType === "to" ? { name, email } : null;
            })
                .filter((entry) => entry !== null), { name: "undisclosed-recipients" }),
            cc: recipients
                .map(({ name, email, recipType }) => recipType === "cc" ? { name, email } : null)
                .filter((entry) => entry !== null),
            bcc: recipients
                .map(({ name, email, recipType }) => recipType === "bcc" ? { name, email } : null)
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
                });
            }
            else {
                const exported = this._msgReader.getAttachment(attachment);
                attachmentsRefined.push({
                    filename: exported.fileName,
                    content: buffer_1.Buffer.from(exported.content),
                    cid: attachment.pidContentId,
                });
            }
        }
        return await new Promise((resolve, reject) => {
            try {
                const mail = new mail_composer_1.default(entity);
                mail.compile().build(function (error, message) {
                    if (error) {
                        reject(new Error("EML composition failed.\n" + error + "\n\n" + error.stack));
                        return;
                    }
                    else {
                        resolve(message);
                    }
                });
            }
            catch (ex) {
                reject(new Error("EML composition failed.\n" + ex));
            }
        });
    }
}
exports.ParsedMsg = ParsedMsg;
class MsgConverter {
    constructor(options) {
        this.options = options || {};
    }
    /**
     * Parse at first
     *
     * @param input msg file
     * @returns parsed entity
     */
    async parse(input) {
        const inputMsg = new msgreader_1.default(input);
        inputMsg.parserConfig = {
            ansiEncoding: this.options.ansiEncoding,
        };
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
    async convertToString(input) {
        return (await this.parse(input)).toEmlStr();
    }
    /**
     * Get EML
     *
     * @param input msg file
     * @returns EML file
     */
    async convertToBuffer(input) {
        return (await this.parse(input)).toEmlBuffer();
    }
}
exports.MsgConverter = MsgConverter;
