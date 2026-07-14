import { describe, expect, test } from "bun:test";

import { serializeFontTableXml } from "./fontTableSerializer";

describe("serializeFontTableXml", () => {
  test("omits undefined font signature attributes", () => {
    const xml = serializeFontTableXml({
      fonts: [
        {
          name: "Arial",
          sig: {
            usb0: "00000000",
            usb1: undefined,
          },
        },
      ],
    });

    expect(xml).toContain('<w:sig w:usb0="00000000"/>');
    expect(xml).not.toContain("w:usb1");
    expect(xml).not.toContain("undefined");
  });
});
