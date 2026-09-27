import { expect, test } from "bun:test";

import { countOpaqueRevisionWrappers, opaqueRevisionCarrierName } from "./opaqueCarrier";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

test("opaque revision classification uses expanded names and counts nested carriers", () => {
  const xml =
    `<w:tbl xmlns:w="${WORDPROCESSINGML_NAMESPACE}" xmlns:x="urn:example:foreign">` +
    "<x:ins><w:del/><w:moveFrom><w:moveTo/></w:moveFrom></x:ins>" +
    "</w:tbl>";

  expect(opaqueRevisionCarrierName(xml)).toBe("w:del");
  expect(countOpaqueRevisionWrappers(xml)).toBe(3);
  expect(opaqueRevisionCarrierName(`<x:ins xmlns:x="urn:example:foreign"/>`)).toBeUndefined();
  expect(countOpaqueRevisionWrappers(`<x:del xmlns:x="urn:example:foreign"/>`)).toBe(0);
});
