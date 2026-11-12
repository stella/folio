import { test } from "@playwright/test";

import {
  PROTOCOL_NAMESPACE,
  playwrightDebug,
} from "../../../tests/support/playwrightProtocolCapture";

const count = (name: string) => {
  const value = Number(process.env[name]);
  if (!Number.isSafeInteger(value) || value < 1)
    throw new TypeError(`${name} must be a positive integer`);
  return value;
};

test("protocol logging floods the worker", () => {
  const log = playwrightDebug()(PROTOCOL_NAMESPACE);
  const expression = "x".repeat(count("FOLIO_PROTOCOL_FLOOD_PAYLOAD_BYTES"));
  const messages = count("FOLIO_PROTOCOL_FLOOD_MESSAGES");
  for (let id = 1; id <= messages; id++)
    log(
      `SEND ► ${JSON.stringify({ id, sessionId: "flood", method: "Runtime.evaluate", params: { expression } })}`,
    );
});
