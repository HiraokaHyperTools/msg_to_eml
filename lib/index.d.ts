/// <reference types="node" />
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
export declare class MsgConverter {
    options: MsgConverterOptions;
    constructor(options: MsgConverterOptions);
    convertToString(input: Uint8Array): Promise<string>;
    convertToBuffer(input: Uint8Array): Promise<Buffer>;
}
