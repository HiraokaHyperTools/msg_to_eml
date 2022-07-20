const assert = require('assert');
const fs = require('fs');
const path = require('path');

const generate = false;

describe('MsgConverter', function () {
  const { MsgConverter } = require('../lib/index');

  it('not null', function () {
    assert.notEqual(MsgConverter, null);
    assert.notEqual(new MsgConverter(), null);
  });

  describe('Conversion exact match', function () {
    const converter = new MsgConverter({
      baseBoundary: "42fb9c3c-c8ed-4b9b-aebf-8379cb3000d5",
      messageId: "cea35f87-010e-4e67-b059-9bb20b55b27c",
      ansiEncoding: "cp932",
    });

    function consumeBad(baseName) {
      return async function () {
        const msgBuf = await fs.promises.readFile(path.join(__dirname, `${baseName}.msg`));
        try {
          await converter.convertToString(msgBuf);
          assert.fail('must fail');
        }
        catch (ex) {
          assert.throws(() => { throw ex }, new Error("MSG file parser error: Unsupported file type!"));
        }
      }
    }

    function consume(baseName) {
      return async function () {
        const msgBuf = await fs.promises.readFile(path.join(__dirname, `${baseName}.msg`));
        const emlStrActual = await converter.convertToString(msgBuf);
        if (generate) {
          fs.promises.writeFile(path.join(__dirname, `${baseName}.eml`), emlStrActual);
        }
        else {
          const emlStrExpected = await fs.promises.readFile(
            path.join(__dirname, `${baseName}.eml`),
            { encoding: 'utf-8' }
          );
          assert.equal(emlStrActual, emlStrExpected);
        }
      }
    }

    it("dummy (bad data test)", consumeBad("dummy"));
    it("msgInMsg", consume("msgInMsg"));
    it("msgInMsgInMsg", consume("msgInMsgInMsg"));
    it("nonUnicodeCP932", consume("nonUnicodeCP932"));
    it("nonUnicodeMail", consume("nonUnicodeMail"));
    it("test1", consume("test1"));
    it("unicode1", consume("unicode1"));
  });
});
