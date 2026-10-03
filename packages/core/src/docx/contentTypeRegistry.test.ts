import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { registerContentTypeParts } from "./contentTypeRegistry";

const namespace = "http://schemas.openxmlformats.org/package/2006/content-types";
const footerType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

test(
  "content-type registration preserves authored entries and treats equivalent part URIs once",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom("", "ct:", "pkg:"),
        fc.constantFrom('"', "'"),
        fc.constantFrom("xml", "XML"),
        fc.boolean(),
        fc.boolean(),
        (prefix, quote, extension, useDefault, encodeName) => {
          const path = "word/článek 日本.xml";
          const uri = encodeName ? "/word/%C4%8Dl%C3%A1nek%20%E6%97%A5%E6%9C%AC.xml" : `/${path}`;
          const binding = prefix ? `xmlns:${prefix.slice(0, -1)}` : "xmlns";
          const entry = useDefault
            ? `<${prefix}Default ContentType=${quote}${footerType}${quote} Extension=${quote}${extension}${quote}/>`
            : `<${prefix}Override ContentType=${quote}${footerType}${quote} PartName=${quote}${uri}${quote}/>`;
          const source = `<?xml version="1.0"?>\r\n<${prefix}Types ${binding}=${quote}${namespace}${quote}>\r\n  ${entry}\n  <!-- untouched -->\n  <alien:Override xmlns:alien="urn:foreign" PartName="/unrelated.xml" ContentType="untouched"/>\n</${prefix}Types>\n<!-- </${prefix}Types> trailing -->`;
          const own = { partName: `/${path}`, contentType: footerType };
          expect(registerContentTypeParts(source, [own])).toBe(source);
          const added = { partName: "/word/new.bin", contentType: "application/octet-stream" };
          const next = registerContentTypeParts(source, [added]);
          const insertion = `<${prefix}Override PartName="${added.partName}" ContentType="${added.contentType}"/>`;
          expect(next.replace(insertion, "")).toBe(source);
          expect(registerContentTypeParts(next, [own, added])).toBe(next);
        },
      ),
      { numRuns: 60 },
    );
  },
  propertyTestTimeout(30_000),
);

test("a foreign same-named entry cannot suppress the actual part registration", () => {
  const source = `<Types xmlns="${namespace}"><x:Override xmlns:x="urn:foreign" PartName="/word/footer.xml" ContentType="wrong"/></Types>`;
  const part = { partName: "/word/footer.xml", contentType: footerType };
  const updated = registerContentTypeParts(source, [part]);
  expect(updated).toBe(
    source.replace(
      "</Types>",
      `<Override PartName="/word/footer.xml" ContentType="${footerType}"/></Types>`,
    ),
  );
});

test("empty root registration preserves quoted boundaries and trailing comments", () => {
  const source = `<Types xmlns="${namespace}" data-extra="/>"/> <!-- /> </Types> -->`;
  const updated = registerContentTypeParts(source, [
    { partName: "/word/footer.xml", contentType: footerType },
  ]);
  expect(updated).toBe(
    `<Types xmlns="${namespace}" data-extra="/>"><Override PartName="/word/footer.xml" ContentType="${footerType}"/></Types> <!-- /> </Types> -->`,
  );
});

test("an extension default for another type permits the owned part override", () => {
  const source = `<Types xmlns="${namespace}"><Default Extension="xml" ContentType="application/xml"/></Types>`;
  const part = { partName: "/word/footer.xml", contentType: footerType };
  const updated = registerContentTypeParts(source, [part]);
  expect(
    updated.replace(`<Override PartName="${part.partName}" ContentType="${footerType}"/>`, ""),
  ).toBe(source);
  expect(registerContentTypeParts(updated, [part])).toBe(updated);
});
