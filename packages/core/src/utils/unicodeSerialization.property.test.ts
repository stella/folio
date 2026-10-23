import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Window } from "happy-dom";
import * as Y from "yjs";

import {
  compileMarkdownToContent,
  serializeDocumentToDocx,
} from "@stll/docx-core";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  OP_STORIES,
  paragraphLogicalText,
  parseDocumentBatch,
} from "@stll/docx-core/ops";
import {
  assertTextRoundTrip,
  astralText,
  boundaryTextCases,
} from "@stll/utf8-boundaries/testing";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { extractDocxText } from "../docx/server/extractDocxText";
import {
  decodeAutoSaveEnvelope,
  encodeAutoSaveEnvelope,
} from "../managers/autoSaveCodec";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toMarkdown } from "../markdown/index";
import { writePdf } from "../pdf/writePdf";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import {
  paragraphsToClipboardContent,
  parseClipboardHtml,
  runsToClipboardContent,
} from "./clipboard";
import {
  bytesToHex,
  createTextReader,
  decodeText,
  encodeUtf8,
  readBlobText,
} from "./textEncoding";

setDefaultTimeout(propertyTestTimeout(180_000));

const timestamp = "2026-01-01T00:00:00.000Z";
const clipboardWindow = new Window();
const parseClipboardText = (html: string, plainText: string): string => {
  const globals = {
    window: clipboardWindow,
    DOMParser: clipboardWindow.DOMParser,
    Node: clipboardWindow.Node,
    HTMLElement: clipboardWindow.HTMLElement,
  };
  const previous = new Map(
    Object.keys(globals).map((key) => [
      key,
      Object.getOwnPropertyDescriptor(globalThis, key),
    ]),
  );
  for (const [key, value] of Object.entries(globals))
    Object.defineProperty(globalThis, key, {
      configurable: true,
      writable: true,
      value,
    });
  try {
    return parseClipboardHtml(html, plainText)
      .runs.flatMap((run) =>
        run.content.flatMap((node) =>
          node.type === "text" ? [node.text] : [],
        ),
      )
      .join("");
  } finally {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
};
const documentText = ({
  package: {
    document: { content },
  },
}: Document): string =>
  content
    .flatMap((block) =>
      block.type === "paragraph" ? [paragraphLogicalText(block)] : [],
    )
    .join("\n");

const roundTripText = (text: string, chunkSize: number): string => {
  const bytes = encodeUtf8(text);
  const reader = createTextReader({ fatal: true, ignoreBOM: true });
  const parts: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    parts.push(reader.read(bytes.subarray(offset, offset + chunkSize)));
  }
  const restored = parts.join("") + reader.finish();
  expect(encodeUtf8(restored)).toEqual(bytes);
  return restored;
};

type SerializerCase = (
  text: string,
  document: Document,
  paragraph: Paragraph,
  chunkSize: number,
) => Promise<void>;

// Each entry exercises its real producer and reader, rather than a substitute codec.
const serializers = {
  docxModel: async (text, document) => {
    const bytes = await serializeDocumentToDocx(document);
    const restored = await parseDocx(bytes, { preloadFonts: false });
    expect(encodeUtf8(documentText(restored))).toEqual(encodeUtf8(text));
    const extracted = await extractDocxText(bytes);
    expect(
      extracted.paragraphs
        .filter((paragraph) => paragraph.source === "body")
        .map((paragraph) => paragraph.text)
        .join("\n"),
    ).toBe(text);
  },
  docxEditor: async (text, document) => {
    const bytes = await createDocx(document);
    const restored = await parseDocx(bytes, { preloadFonts: false });
    expect(encodeUtf8(documentText(restored))).toEqual(encodeUtf8(text));
    const savedAgain = await createDocx(restored);
    expect(
      documentText(await parseDocx(savedAgain, { preloadFonts: false })),
    ).toBe(text);
  },
  snapshot: async (text) => {
    const node = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "00000001" }, schema.text(text)),
    ]);
    const snapshot = createFolioAIEditSnapshot(node);
    expect(snapshot.blocks.map((block) => block.text).join("\n")).toBe(text);
  },
  markdown: async (text, document, _paragraph, chunkSize) => {
    const markdown = roundTripText(toMarkdown(document), chunkSize);
    expect(documentText(fromMarkdown(markdown))).toBe(text);
    const { content } = compileMarkdownToContent(markdown);
    expect(
      content
        .flatMap((block) =>
          block.type === "paragraph" ? [paragraphLogicalText(block)] : [],
        )
        .join("\n"),
    ).toBe(text);
  },
  clipboard: async (text, _document, paragraph, chunkSize) => {
    const runs = paragraph.content.flatMap((node) =>
      node.type === "run" ? [node] : [],
    );
    for (const { content, expected } of [
      { content: runsToClipboardContent(runs), expected: runs },
      {
        content: paragraphsToClipboardContent([paragraph]),
        expected: [paragraph],
      },
    ]) {
      expect(roundTripText(content.plainText, chunkSize)).toBe(text);
      expect(roundTripText(content.html, chunkSize)).toBe(content.html);
      expect(
        parseClipboardText(roundTripText(content.html, chunkSize), ""),
      ).toBe(text);
      expect(
        await readBlobText(new Blob([encodeUtf8(content.plainText)])),
      ).toBe(text);
      expect(await readBlobText(new Blob([encodeUtf8(content.html)]))).toBe(
        content.html,
      );
      const internal = roundTripText(content.internal ?? "", chunkSize);
      expect(JSON.parse(internal)).toEqual(expected);
      // The internal JSON must contain the original text, not a replacement character.
      expect(internal).toContain(text);
    }
  },
  autosave: async (text, document, _paragraph, chunkSize) => {
    const wire = roundTripText(
      encodeAutoSaveEnvelope(document, timestamp),
      chunkSize,
    );
    const restored = decodeAutoSaveEnvelope(wire);
    expect(restored).not.toBeNull();
    if (!restored)
      throw new TypeError("Autosave rejected its own serialized model");
    expect(documentText(restored.document)).toBe(text);
  },
  batch: async (text, _document, _paragraph, chunkSize) => {
    const batch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: "01JUNICODE0000000000000001",
      actor: "editor",
      baseRev: 0,
      ops: [
        {
          type: "insertText",
          at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
          text,
          runProps: "inherit",
        },
      ],
    };
    const wire = roundTripText(JSON.stringify(batch), chunkSize);
    const decoded = parseDocumentBatch(wire);
    expect(decoded.isOk()).toBe(true);
    if (decoded.isErr()) throw decoded.error;
    expect(encodeUtf8(JSON.stringify(decoded.value))).toEqual(encodeUtf8(wire));
  },
  collaboration: async (text, _document, _paragraph, chunkSize) => {
    const source = new Y.Doc();
    const target = new Y.Doc();
    source.getText("body").insert(0, text);
    const update = Y.encodeStateAsUpdate(source);
    const transported = new Uint8Array(update.length);
    for (let offset = 0; offset < update.length; offset += chunkSize)
      transported.set(update.subarray(offset, offset + chunkSize), offset);
    expect(transported).toEqual(update);
    Y.applyUpdate(target, transported);
    expect(encodeUtf8(target.getText("body").toString())).toEqual(
      encodeUtf8(text),
    );
    source.destroy();
    target.destroy();
  },
  pdf: async (text, _document, _paragraph, chunkSize) => {
    const result = await writePdf(
      {
        pages: [
          {
            pageNumber: 1,
            widthPx: 816,
            heightPx: 1056,
            orientation: "portrait",
            regions: [],
            primitives: [],
            links: [],
            comments: [],
          },
        ],
        fonts: [],
        images: [],
        outline: [],
        metadata: { title: text },
        unsupported: [],
      },
      { timestamp, fonts: { load: () => [] } },
    );
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    const bytes = result.value.bytes;
    const transported = new Uint8Array(bytes.length);
    for (let offset = 0; offset < bytes.length; offset += chunkSize)
      transported.set(bytes.subarray(offset, offset + chunkSize), offset);
    expect(transported).toEqual(bytes);
    const title = decodeText(bytes).match(/\/Title <FEFF(?<hex>[0-9A-F]*)>/u)
      ?.groups?.["hex"];
    expect(title).toBeDefined();
    const titleBytes = Uint8Array.from(Buffer.from(title ?? "", "hex"));
    expect(decodeText(titleBytes, { encoding: "utf-16be", fatal: true })).toBe(
      text,
    );
    expect(bytesToHex(titleBytes)).toBe((title ?? "").toLowerCase());
  },
} satisfies Record<string, SerializerCase>;

// Bound each test's work without reducing the census, generated runs, or pinned replays.
const MAX_BOUNDARY_CASES_PER_TEST = 5;
const boundaryCaseCount = boundaryTextCases("", "").length;
const caseStarts = Array.from(
  { length: boundaryCaseCount },
  (_, index) => index,
).filter((index) => index % MAX_BOUNDARY_CASES_PER_TEST === 0);

// Keep the shared property ID: every producer and case group replays the pinned inputs.
for (const [name, serializer] of Object.entries(serializers)) {
  for (const caseStart of caseStarts) {
    test(`${name} retains Unicode across byte and buffer boundaries (cases ${caseStart})`, async () => {
      await assertProperty(
        fc.asyncProperty(
          astralText,
          fc.array(astralText, { maxLength: 4 }),
          async (astral, suffixPoints) => {
            const cases = boundaryTextCases(astral, suffixPoints.join(""));
            expect(cases).toHaveLength(boundaryCaseCount);
            for (const { boundary, text } of cases.slice(
              caseStart,
              caseStart + MAX_BOUNDARY_CASES_PER_TEST,
            )) {
              const paragraph: Paragraph = {
                type: "paragraph",
                paraId: "00000001",
                content: [{ type: "run", content: [{ type: "text", text }] }],
              };
              const document: Document = {
                package: { document: { content: [paragraph] } },
              };
              await assertTextRoundTrip({
                text,
                roundTrip: (value) => roundTripText(value, boundary),
              });
              await serializer(text, document, paragraph, boundary);
            }
          },
        ),
        {
          numRuns: 3,
          id: "document serializers retain Unicode across byte boundaries",
        },
      );
    });
  }
}
