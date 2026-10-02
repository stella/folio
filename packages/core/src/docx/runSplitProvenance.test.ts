import { describe, expect, test } from "bun:test";
import type { Document, Run } from "../types/document";
import { expectRunIdentityMarkAttrs } from "../prosemirror/attrs";
import { RUN_IDENTITY_MARK_NAME } from "../prosemirror/runIdentity";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { serializeDocument } from "./serializer/documentSerializer";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import {
  InvalidRunSplitProvenanceError,
  readRunSplitOwner,
  withRunSplitOwner,
} from "./runSplitProvenance";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });
const documentOf = (runs: Run[]): Document => ({
  package: { document: { content: [{ type: "paragraph", content: runs }] } },
});
const owners = (document: Document) => {
  const result: ReturnType<typeof expectRunIdentityMarkAttrs>[] = [];
  toProseDoc(document).descendants((node) => {
    if (!node.isText) return;
    const mark = node.marks.find(({ type }) => type.name === RUN_IDENTITY_MARK_NAME);
    if (mark === undefined) throw new TypeError("Expected projected source-run identity");
    result.push(expectRunIdentityMarkAttrs(mark));
  });
  return result;
};

describe("explicit saved split-run ownership", () => {
  test("only validated proof in the declared namespace aliases imported owners", () => {
    const proof = "1:main:1:0";
    const left = { ...run("left"), preservedAttributes: withRunSplitOwner(undefined, proof) };
    const right = {
      ...run("right"),
      formatting: { bold: true },
      preservedAttributes: withRunSplitOwner(undefined, proof),
    };
    const same = owners(documentOf([left, right]));
    expect(same.at(0)?.id).toBe(same.at(1)?.id);
    const foreign = [left, right].map((value) =>
      Object.assign({}, value, {
        preservedAttributes: [
          { namespace: OOXML_NAMESPACES.w.uri, name: "splitRunOwner", value: proof },
        ],
      }),
    );
    expect(readRunSplitOwner(foreign.at(0)?.preservedAttributes)).toBeUndefined();
    const distinct = owners(documentOf(foreign));
    expect(distinct.at(0)?.id).not.toBe(distinct.at(1)?.id);
  });

  test("malformed, duplicate and exhausted proof refuse rather than aliasing", () => {
    for (const value of [
      "",
      "0:main:1:0",
      "1:main:-1:0",
      "1:main:1:NaN",
      "1:main:9007199254740992:0",
    ]) {
      const attributes = withRunSplitOwner(undefined, value);
      expect(() =>
        owners(documentOf([{ ...run("text"), preservedAttributes: attributes }])),
      ).toThrow(InvalidRunSplitProvenanceError);
    }
    const attributes = withRunSplitOwner(undefined, "1:main:1:0");
    expect(() => readRunSplitOwner([...attributes, ...attributes])).toThrow(
      InvalidRunSplitProvenanceError,
    );
    expect(() =>
      owners(
        documentOf([
          {
            ...run("text"),
            preservedAttributes: withRunSplitOwner(undefined, "1:main:9007199254740991:0"),
          },
        ]),
      ),
    ).toThrow(InvalidRunSplitProvenanceError);
  });

  test("newly authored owners cannot reuse a loaded generation after run ordering changes", () => {
    const existing = {
      ...run("existing"),
      preservedAttributes: withRunSplitOwner(undefined, "1:main:1:0"),
    };
    for (const content of [
      [run("new"), existing],
      [existing, run("new")],
    ]) {
      const projected = owners(documentOf(content));
      expect(projected.at(0)?.id).not.toBe(projected.at(1)?.id);
      const proofs = projected.map(({ preservedAttributes }) =>
        readRunSplitOwner(preservedAttributes),
      );
      expect(new Set(proofs).size).toBe(2);
      expect(proofs).toContain("1:main:1:0");
      expect(proofs.some((proof) => proof?.startsWith("1:main:2:"))).toBe(true);
      expect(owners(documentOf(content))).toEqual(projected);
    }
  });

  test("saved provenance has a bound namespace and is explicitly ignorable", () => {
    const document = documentOf([
      {
        ...run("text"),
        preservedAttributes: withRunSplitOwner(undefined, "1:main:1:0"),
      },
    ]);
    const xml = serializeDocument(document);
    expect(xml).toContain(`xmlns:folio="${OOXML_NAMESPACES.folio.uri}"`);
    expect(xml).toMatch(/mc:Ignorable="[^"]*\bfolio\b[^"]*"/u);
    expect(xml).toContain('folio:splitRunOwner="1:main:1:0"');
  });
});
