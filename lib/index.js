"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MsgConverter = void 0;
const msgreader_1 = __importDefault(require("@kenjiuno/msgreader"));
const decompressrtf_1 = require("@kenjiuno/decompressrtf");
const iconv_lite_1 = __importDefault(require("iconv-lite"));
const rtf_stream_parser_1 = require("rtf-stream-parser");
const mail_composer_1 = __importDefault(require("nodemailer/lib/mail-composer"));
const buffer_1 = require("buffer");
const path_1 = __importDefault(require("path"));
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
class MsgConverter {
    constructor(options) {
        this.options = options || {};
    }
    async convertToString(input) {
        return (await this.convertToBuffer(input)).toString('utf-8');
    }
    async convertToBuffer(input) {
        const inputMsg = new msgreader_1.default(input);
        inputMsg.parserConfig = {
            ansiEncoding: this.options.ansiEncoding,
        };
        const msgInfo = inputMsg.getFileData();
        if (msgInfo.error) {
            throw new Error(`MSG file parser error: ${msgInfo.error}`);
        }
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
        for (let attachment of attachments) {
            const exported = inputMsg.getAttachment(attachment);
            if (attachment.innerMsgContent) {
                const emlBuf = await this.convertToBuffer(exported.content);
                attachmentsRefined.push({
                    filename: changeExtension(exported.fileName, ".eml"),
                    content: emlBuf,
                    cid: attachment.pidContentId,
                    contentTransferEncoding: '8bit',
                });
            }
            else {
                attachmentsRefined.push({
                    filename: exported.fileName,
                    content: buffer_1.Buffer.from(exported.content),
                    cid: attachment.pidContentId,
                });
            }
        }
        const entity = {
            baseBoundary: this.options.baseBoundary,
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
                "Message-ID": this.options.messageId,
            }
        };
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
exports.MsgConverter = MsgConverter;
function changeExtension(fileName, newExt) {
    const parsed = path_1.default.parse(fileName);
    return (parsed.dir ? parsed.dir + path_1.default.sep : "") + parsed.name + newExt;
}
