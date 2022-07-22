const program = require('commander');

const { MsgConverter } = require('./lib/index');

const fs = require('fs');

program
  .command('convert <msgFilePath> [saveEMLTo]')
  .description('Convert msg file and print eml')
  .option('--ansi-encoding <encoding>', 'Set ANSI encoding (used by iconv-lite) for non Unicode text in msg file')
  .action(async (msgFilePath, saveEMLTo, options) => {
    try {
      const msgFileBuffer = await fs.promises.readFile(msgFilePath);
      const buffer = await new MsgConverter({ ansiEncoding: options.ansiEncoding }).convertToBuffer(msgFileBuffer);

      if (saveEMLTo) {
        await fs.promises.writeFile(saveEMLTo, buffer);
      }
      else {
        process.stdout.write(buffer);
      }
    } catch (ex) {
      process.exitCode = 1;
      console.error(ex);
    }
  });

program
  .command('vcard <msgFilePath> [saveVCardTo]')
  .description('Convert msg file and print vCard')
  .option('--ansi-encoding <encoding>', 'Set ANSI encoding (used by iconv-lite) for non Unicode text in msg file')
  .action(async (msgFilePath, saveVCardTo, options) => {
    try {
      const msgFileBuffer = await fs.promises.readFile(msgFilePath);
      const buffer = Buffer.from(
        await (
          await new MsgConverter({ ansiEncoding: options.ansiEncoding })
            .parse(msgFileBuffer)
        )
          .toVCardStr()
      );

      if (saveVCardTo) {
        await fs.promises.writeFile(saveVCardTo, buffer);
      }
      else {
        process.stdout.write(buffer);
      }
    } catch (ex) {
      process.exitCode = 1;
      console.error(ex);
    }
  });

program
  .parse(process.argv);
