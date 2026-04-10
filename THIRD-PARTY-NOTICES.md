# Third-Party Notices

This document lists third-party software and reference materials used or referred to by `mikuproject`.

## Third-party software

### `jsdom`

- Usage: Used in the CLI and tests to provide browser-like Web APIs such as `window`, `document`, `Blob`, and `File`. XML processing in the CLI no longer depends on `jsdom` directly, but `jsdom` remains in use for non-XML Web API emulation.
- License: MIT
- Source: https://github.com/jsdom/jsdom

### `@xmldom/xmldom`

- Usage: Used in the CLI as the preferred XML DOM implementation for `DOMParser`, `XMLSerializer`, and XML document creation.
- License: MIT
- Source: https://github.com/xmldom/xmldom

## Reference materials

### `open-msp-viewer`

- Usage: Referred to as a temporary source of sample data for verification
- License: Apache License 2.0
- Source: https://github.com/rpbouman/open-msp-viewer/

### MicrosoftDocs Project XML Data Interchange reference

- Usage: Referred to when design decisions for `MS Project XML` handling are unclear. This reference is used together with the Microsoft-hosted Project 2007 schema endpoints such as `https://schemas.microsoft.com/project/2007/` and `mspdi_pj12.xsd`.
- License: CC-BY-4.0 for documentation and MIT for code
- Source: https://github.com/MicrosoftDocs/office-developer-msproject-xml-docs/tree/main/project-xml-data-interchange

---

# 第三者告知

この文書は、`mikuproject` が利用または参照している第三者ソフトウェアおよび参考資料を記載したものです。

## 第三者ソフトウェア

### `jsdom`

- 用途: CLI とテストで、`window`、`document`、`Blob`、`File` などのブラウザ相当 Web API を補うために利用する。CLI の XML 処理は `jsdom` 直依存ではなくなったが、XML 以外の Web API エミュレーション用途では引き続き利用する。
- ライセンス: MIT
- Source: https://github.com/jsdom/jsdom

### `@xmldom/xmldom`

- 用途: CLI で、`DOMParser`、`XMLSerializer`、XML document 生成の優先 XML DOM 実装として利用する。
- ライセンス: MIT
- Source: https://github.com/xmldom/xmldom

## 参考資料

### `open-msp-viewer`

- 用途: 検証用サンプルデータの一時的な参照元
- ライセンス: Apache License 2.0
- Source: https://github.com/rpbouman/open-msp-viewer/

### MicrosoftDocs Project XML Data Interchange reference

- 用途: `MS Project XML` の扱いで設計判断に迷った場合の補助資料。`https://schemas.microsoft.com/project/2007/` や `mspdi_pj12.xsd` などの Microsoft 側 schema 実体とあわせて参照する。
- ライセンス: 文書は CC-BY-4.0、コードは MIT
- Source: https://github.com/MicrosoftDocs/office-developer-msproject-xml-docs/tree/main/project-xml-data-interchange
