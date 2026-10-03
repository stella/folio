import type { DocumentBody } from "../../src/model/document";

export const assertReadonlySourceXml = (body: DocumentBody) => {
  if (body.source === undefined) return;
  // @ts-expect-error Captured source XML is immutable.
  body.source.xml = "changed";
};
