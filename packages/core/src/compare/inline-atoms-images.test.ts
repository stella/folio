import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import { prepareTargetInlineAtom } from "./inline-atom-resources";
import { matchInlineAtoms, sameInlineAtoms } from "./inline-atoms";

const BASE_IMAGE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
const TARGET_IMAGE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

const image = ({ src, rId }: { src: string; rId: string }) =>
  schema.nodes["image"]!.create({
    src,
    rId,
    width: 96,
    height: 96,
    wrapType: "inline",
    displayMode: "inline",
  });

const rawDrawing = ({ id, rId, name }: { id: string; rId: string; name: string }) =>
  `<w:drawing><wp:inline><wp:docPr id="${id}" name="${name}"/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image.png"/>` +
  `</pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill>` +
  "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";

const imageWithDrawing = ({
  id,
  rId,
  name,
  src = BASE_IMAGE,
}: {
  id: string;
  rId: string;
  name: string;
  src?: string;
}) => {
  const node = image({ src, rId });
  return node.type.create({
    ...node.attrs,
    docPrId: id,
    _docxRawXml: rawDrawing({ id, rId, name }),
  });
};

const field = () =>
  schema.nodes["field"]!.create({
    fieldType: "NUMPAGES",
    instruction: " NUMPAGES ",
    displayText: "1",
    fieldKind: "simple",
  });

const documentWith = (content: readonly PMNode[]) =>
  schema.node("doc", null, [schema.node("paragraph", null, content)]);

const resolve = ({
  state,
  mode,
}: {
  state: EditorState;
  mode: "accept" | "reject";
}): EditorState => {
  let resolved = state;
  const command = mode === "accept" ? acceptAllChanges() : rejectAllChanges();
  command(state, (transaction) => {
    resolved = state.apply(transaction);
  });
  return resolved;
};

const match = ({ state, target }: { state: EditorState; target: PMNode }) =>
  matchInlineAtoms({
    state,
    targetSnapshot: createFolioAIEditSnapshot(target),
    revisionStamp: { idSeed: 40, date: "2026-09-13T00:00:00.000Z" },
    originalRevisionIdSeed: 10,
    author: "Compare",
    maxRanges: 10,
  });

const inlineNames = (document: PMNode): string[] => {
  const paragraph = document.firstChild;
  if (!paragraph) {
    throw new Error("Expected paragraph");
  }
  const names: string[] = [];
  paragraph.forEach((node) => names.push(node.type.name));
  return names;
};

describe("matchInlineAtoms image resources", () => {
  test("compares image content across regenerated drawing and relationship ids", () => {
    const source = documentWith([
      imageWithDrawing({ id: "1", rId: "rIdOriginal", name: "Picture" }),
    ]);
    const renumbered = documentWith([
      imageWithDrawing({ id: "100000", rId: "rIdRepacked", name: "Picture" }),
    ]);
    const renamed = documentWith([
      imageWithDrawing({ id: "100000", rId: "rIdRepacked", name: "Other picture" }),
    ]);
    const differentMedia = documentWith([
      imageWithDrawing({ id: "100000", rId: "rIdRepacked", name: "Picture", src: TARGET_IMAGE }),
    ]);

    expect(sameInlineAtoms(source, renumbered)).toBe(true);
    expect(sameInlineAtoms(source, renamed)).toBe(false);
    expect(sameInlineAtoms(source, differentMedia)).toBe(false);
  });
  test("does not track an unchanged image", () => {
    const unchanged = image({ src: BASE_IMAGE, rId: "rIdBase" });
    const target = documentWith([unchanged]);
    const state = EditorState.create({ schema, doc: documentWith([unchanged]) });

    const result = match({ state, target });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    expect(result.rangeCount).toBe(0);
    expect(result.transaction.doc.eq(state.doc)).toBe(true);
  });

  test("keeps original image media on reject and detached target media on accept", () => {
    const original = image({ src: BASE_IMAGE, rId: "rIdBase" });
    const targetImage = image({ src: TARGET_IMAGE, rId: "rIdTarget" });
    const state = EditorState.create({ schema, doc: documentWith([original]) });
    const result = match({ state, target: documentWith([targetImage]) });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);
    const acceptedImage = resolve({ state: reviewed, mode: "accept" }).doc.firstChild!.firstChild!;
    const rejectedImage = resolve({ state: reviewed, mode: "reject" }).doc.firstChild!.firstChild!;
    const preparedTarget = prepareTargetInlineAtom(targetImage);

    expect(result.rangeCount).toBe(2);
    expect(rejectedImage.attrs["src"]).toBe(BASE_IMAGE);
    expect(rejectedImage.attrs["rId"]).toBe("rIdBase");
    expect(acceptedImage.attrs["src"]).toBe(TARGET_IMAGE);
    expect(acceptedImage.attrs["rId"]).toBe(preparedTarget?.attrs["rId"]);
  });

  test("reuses an unchanged field beside a changed image at one text offset", () => {
    const originalImage = image({ src: BASE_IMAGE, rId: "rIdBase" });
    const targetImage = image({ src: TARGET_IMAGE, rId: "rIdTarget" });
    const source = documentWith([originalImage, field()]);
    const target = documentWith([field(), targetImage]);
    const state = EditorState.create({ schema, doc: source });
    const result = match({ state, target });

    expect(result.status).toBe("matched");
    if (result.status !== "matched") return;
    const reviewed = state.apply(result.transaction);

    expect(result.rangeCount).toBe(2);
    const preservedField = reviewed.doc.firstChild?.child(1);
    expect(preservedField?.type.name).toBe("field");
    expect(preservedField?.marks).toHaveLength(0);
    expect(inlineNames(resolve({ state: reviewed, mode: "reject" }).doc)).toEqual([
      "image",
      "field",
    ]);
    expect(inlineNames(resolve({ state: reviewed, mode: "accept" }).doc)).toEqual([
      "field",
      "image",
    ]);
  });
});
