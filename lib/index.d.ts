/// <reference types="node" />
import MsgReader, { FieldsData } from '@kenjiuno/msgreader';
import { Buffer } from 'buffer';
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
export declare class ParsedMsg {
    private _msgReader;
    private _msgInfo;
    private _options;
    constructor(msgReader: MsgReader, msgInfo: FieldsData, options: MsgConverterOptions);
    get msgReader(): MsgReader;
    get msgInfo(): FieldsData;
    /**
     * Decide convertibility delivered from message class.
     *
     * - vCard, if `IPM.Contact`
     * - EML, if `IPM.Note`
     */
    get messageClass(): string;
    /**
     * Assume this is a contact and then convert to vCard.
     *
     * @returns vCard text
     */
    toVCardStr(): Promise<string>;
    /**
     * Assume this is a mail message and then convert to EML.
     *
     * @returns EML text
     */
    toEmlStr(): Promise<string>;
    /**
     * Assume this is a mail message and then convert to EML.
     *
     * @returns EML file
     */
    toEmlBuffer(): Promise<Buffer>;
    private toEmlFrom;
}
export declare class MsgConverter {
    options: MsgConverterOptions;
    constructor(options: MsgConverterOptions);
    /**
     * Parse at first
     *
     * @param input msg file
     * @returns parsed entity
     */
    parse(input: Uint8Array): Promise<ParsedMsg>;
    /**
     * Get EML
     *
     * @param input msg file
     * @returns EML text
     */
    convertToString(input: Uint8Array): Promise<string>;
    /**
     * Get EML
     *
     * @param input msg file
     * @returns EML file
     */
    convertToBuffer(input: Uint8Array): Promise<Buffer>;
}
