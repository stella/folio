# @stll/folio-connector-remarkable

reMarkable cloud connector for folio. List, upload, download, move, rename, and
delete documents on a reMarkable tablet through the cloud API.

reMarkable accepts PDF and EPUB uploads. folio edits `.docx`; convert to PDF in
your host app before pushing if you want tablet markup on a Word document.

## Install

```sh
bun add @stll/folio-connector-remarkable
```

## Authentication

1. Open https://my.remarkable.com/device/browser/connect and copy the one-time
   code.
2. Exchange it for a long-lived device token and persist it.

```ts
import { registerRemarkableDevice } from "@stll/folio-connector-remarkable";

const deviceToken = await registerRemarkableDevice("abcd1234");
// store deviceToken securely
```

For serverless or worker pools, cache a short-lived session token:

```ts
import {
  authRemarkableSession,
  createRemarkableConnectorFromAuth,
} from "@stll/folio-connector-remarkable";

const sessionToken = await authRemarkableSession(deviceToken);
const connector = await createRemarkableConnectorFromAuth({ sessionToken });
```

## Usage

```ts
import { createRemarkableConnectorFromAuth } from "@stll/folio-connector-remarkable";

const connector = await createRemarkableConnectorFromAuth({ deviceToken });

const library = await connector.listItems();
const matters = await connector.listItems({ parentId: "folder-uuid" });

const uploaded = await connector.uploadDocument({
  name: "Settlement.pdf",
  bytes: pdfBytes,
  mimeType: "application/pdf",
  parentId: "folder-uuid",
});

const originalPdf = await connector.downloadDocument(uploaded);
```

## Connector contract

`FolioCloudConnector` is provider-neutral. The reMarkable implementation sets
`providerId` to `"remarkable"` and exposes the underlying `rmapi-js` client as
`connector.api` when you need provider-specific escape hatches.

## Limitations

- Notebook (`.rm`) downloads are not supported; export from the tablet as PDF
  first.
- Downloaded PDFs are the original uploads, not annotated exports with ink.
- The cloud API is reverse-engineered; upstream changes may require connector
  updates.

## Development

```sh
bun --filter @stll/folio-connector-remarkable test
bun --filter @stll/folio-connector-remarkable typecheck
```

Live cloud tests are intentionally omitted. Run manual smoke tests with your own
device token when validating against production.
