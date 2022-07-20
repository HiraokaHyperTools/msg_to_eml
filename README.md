# msg_to_eml

Links: [_typedoc documentation_](https://hiraokahypertools.github.io/msg_to_eml/typedoc/)

Installation

```bat
yarn add @hiraokahypertools/msg_to_eml
```

Usage:

```js
  const { MsgConverter } = require('@hiraokahypertools/msg_to_eml');

  const msgFileBuffer = await fs.promises.readFile(msgFilePath);
  const buffer = await (new MsgConverter({ ansiEncoding: options.ansiEncoding }).convertToBuffer(msgFileBuffer));
  process.stdout.write(buffer);
```
