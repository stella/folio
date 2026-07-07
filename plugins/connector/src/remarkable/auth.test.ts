import { describe, expect, test } from "bun:test";
import { remarkable } from "rmapi-js";

import {
  authRemarkableSession,
  createRemarkableConnectorFromAuth,
} from "./auth";
import { RemarkableConnectorConfigError } from "../errors";

describe("createRemarkableConnectorFromAuth", () => {
  test("requires deviceToken or sessionToken", async () => {
    await expect(createRemarkableConnectorFromAuth({})).rejects.toBeInstanceOf(
      RemarkableConnectorConfigError,
    );
  });

  test("builds a connector from a session token without network I/O", async () => {
    const connector = await createRemarkableConnectorFromAuth({
      sessionToken: "cached-session-token",
    });

    expect(connector.providerId).toBe("remarkable");
    expect(typeof connector.api.listItems).toBe("function");
    expect(typeof connector.api.uploadPdf).toBe("function");
  });

  test("exports auth helpers that delegate to rmapi-js", () => {
    expect(authRemarkableSession).toBeDefined();
    expect(remarkable).toBeDefined();
  });
});
