import { panic, Result } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { getTrackedChangesFromDoc } from "../ai-edits/read";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationBatch,
} from "../document-operations";
import type { Document, Paragraph, Run } from "../types/document";
import type { DeepReadonly } from "../types/documentModel";

const paragraphOf = (
  paraId: string,
  runs: Run[],
  formatting?: Paragraph["formatting"],
): Paragraph => ({
  type: "paragraph",
  paraId,
  ...(formatting === undefined ? {} : { formatting }),
  content: runs,
});

const textRun = (text: string, formatting?: Run["formatting"]): Run => ({
  type: "run",
  ...(formatting === undefined ? {} : { formatting }),
  content: [{ type: "text", text }],
});

const documentOf = (...content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const setup = (
  document: Document = documentOf(
    paragraphOf("12345678", [textRun("alpha beta")]),
    paragraphOf("23456789", [textRun("gamma")]),
  ),
) => {
  const created = createCanonicalSession(document);
  if (created.isErr()) return panic(created.error.message);
  const session = created.value;
  let state = EditorState.create({ doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const result = publishCanonicalProjection({ state, commit, session });
    if (result.isErr()) return false;
    state = result.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({
    undoLimit: Number.POSITIVE_INFINITY,
    session,
    getState: () => state,
    publish,
  });
  const apply = (
    operations: readonly FolioDocumentOperation[],
    extra: Omit<FolioDocumentOperationBatch, "version" | "operations"> = {},
  ) =>
    executor.apply({
      snapshot: createFolioAIEditSnapshot(state.doc),
      batch: { version: 1, mode: "direct", operations, ...extra },
      undefinedReferences: "refuse",
    });
  const paragraph = (paraId: string): Paragraph => {
    const found = session.document.package.document.content.find(
      (block): block is Paragraph =>
        block.type === "paragraph" && block.paraId === paraId,
    );
    return found ?? panic(`No paragraph ${paraId}`);
  };
  const accept = () => {
    const accepted = session
      .prepareResolve(state, {
        revisionIds: getTrackedChangesFromDoc(state.doc).map(({ id }) => id),
        resolution: "accept",
      })
      .unwrap();
    expect(publish(accepted)).toBe(true);
  };
  return { session, apply, paragraph, accept, state: () => state };
};

/** Each run's text with the emphasis it states, empty runs left out. */
const runsOf = (paragraph: DeepReadonly<Paragraph>) =>
  paragraph.content.flatMap((item) => {
    if (item.type !== "run") return [];
    const text = item.content
      .map((content) => {
        if (content.type === "text") return content.text;
        if (content.type === "tab") return "\t";
        if (content.type === "break") return "\n";
        return "";
      })
      .join("");
    if (text === "") return [];
    return [
      {
        text,
        bold: item.formatting?.bold === true,
        italic: item.formatting?.italic === true,
      },
    ];
  });

/** Adjacent runs of the same emphasis read as one. */
const emphasisOf = (paragraph: Paragraph) => {
  const merged: ReturnType<typeof runsOf> = [];
  for (const run of runsOf(paragraph)) {
    const last = merged.at(-1);
    if (
      last !== undefined &&
      last.bold === run.bold &&
      last.italic === run.italic
    )
      last.text += run.text;
    else merged.push({ ...run });
  }
  return merged;
};

describe("canonical public payloads", () => {
  test("a direct replacement writes the words its emphasis markers mark as bold or italic", () => {
    const editor = setup();
    const result = editor.apply([
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "alpha",
        replace: "**Gamma** and ***delta***",
      },
    ]);
    expect(result.skipped).toEqual([]);
    expect(emphasisOf(editor.paragraph("12345678"))).toEqual([
      { text: "Gamma", bold: true, italic: false },
      { text: " and ", bold: false, italic: false },
      { text: "delta", bold: true, italic: true },
      { text: " beta", bold: false, italic: false },
    ]);
  });

  test("emphasis on the words a match already holds is no no-op", () => {
    const editor = setup();
    const result = editor.apply([
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "alpha",
        replace: "**alpha**",
      },
    ]);
    expect(result.skipped).toEqual([]);
    expect(emphasisOf(editor.paragraph("12345678"))).toEqual([
      { text: "alpha", bold: true, italic: false },
      { text: " beta", bold: false, italic: false },
    ]);
  });

  test("a tracked replacement writes its emphasized words plain, as the redline is text only", () => {
    const editor = setup();
    const result = editor.apply(
      [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: "12345678",
          find: "alpha",
          replace: "**gamma**",
        },
      ],
      { mode: "tracked-changes" },
    );
    expect(result.skipped).toEqual([]);
    editor.accept();
    expect(emphasisOf(editor.paragraph("12345678"))).toEqual([
      { text: "gamma beta", bold: false, italic: false },
    ]);
  });

  test("a direct replacement that drops the paragraph's formatting writes the words in a paragraph that states none", () => {
    const editor = setup(
      documentOf(
        paragraphOf("12345678", [textRun("alpha beta", { italic: true })], {
          alignment: "center",
          spaceAfter: 240,
        }),
        paragraphOf("23456789", [textRun("gamma")]),
      ),
    );
    const result = editor.apply([
      {
        id: "replace",
        type: "replaceBlock",
        blockId: "12345678",
        text: "**New** words",
        preserveFormatting: false,
      },
    ]);
    expect(result.skipped).toEqual([]);
    const paragraph = editor.paragraph("12345678");
    expect(paragraph.formatting?.alignment).toBeUndefined();
    expect(paragraph.formatting?.spaceAfter).toBeUndefined();
    expect(emphasisOf(paragraph)).toEqual([
      { text: "New", bold: true, italic: false },
      { text: " words", bold: false, italic: false },
    ]);
  });

  test("a tracked replacement that drops the paragraph's formatting keeps it, as the redline is text only", () => {
    const editor = setup(
      documentOf(
        paragraphOf("12345678", [textRun("alpha beta")], {
          alignment: "center",
        }),
        paragraphOf("23456789", [textRun("gamma")]),
      ),
    );
    const result = editor.apply(
      [
        {
          id: "replace",
          type: "replaceBlock",
          blockId: "12345678",
          text: "alpha delta",
          preserveFormatting: false,
        },
      ],
      { mode: "tracked-changes" },
    );
    expect(result.skipped).toEqual([]);
    editor.accept();
    const paragraph = editor.paragraph("12345678");
    expect(paragraph.formatting?.alignment).toBe("center");
    expect(emphasisOf(paragraph)).toEqual([
      { text: "alpha delta", bold: false, italic: false },
    ]);
  });

  test("a join separator writes a tab and a line feed as their elements", () => {
    const editor = setup();
    const result = editor.apply([
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: "12345678",
        separator: "\t-\n",
      },
    ]);
    expect(result.skipped).toEqual([]);
    const blocks = editor.session.document.package.document.content;
    expect(blocks).toHaveLength(1);
    const [merged] = blocks;
    if (merged?.type !== "paragraph")
      return panic("Expected the merged paragraph.");
    const contents = merged.content.flatMap((item) =>
      item.type === "run" ? item.content.map(({ type }) => type) : [],
    );
    expect(contents).toContain("tab");
    expect(contents).toContain("break");
    expect(
      runsOf(merged)
        .map(({ text }) => text)
        .join(""),
    ).toBe("alpha beta\t-\ngamma");
  });

  test("a join separator with a carriage return is refused as the request", () => {
    const editor = setup();
    const result = editor.apply([
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: "12345678",
        separator: "\r",
      },
    ]);
    expect(result.skipped.at(0)?.reason).toBe("unsupportedBlock");
    expect(result.issues.at(0)?.canonicalRefusal).toEqual({
      condition: "invalidRequest",
    });
  });

  test("text that writes a note marker of its own is refused as a protected reference", () => {
    const editor = setup();
    const inserted = editor.apply([
      {
        id: "insert",
        type: "insertAfterBlock",
        blockId: "12345678",
        text: "See [^1].",
      },
    ]);
    expect(inserted.skipped.at(0)?.reason).toBe("protectedReference");
    expect(inserted.issues.at(0)?.canonicalRefusal).toBeUndefined();
    const replaced = editor.apply([
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "beta",
        replace: "beta [^0]",
      },
    ]);
    expect(replaced.skipped.at(0)?.reason).toBe("protectedReference");
  });

  test("a batch the session refuses to stage reports the session's own ground", () => {
    const editor = setup();
    const staging = editor.session.beginVerifiedOperationStaging.bind(
      editor.session,
    );
    const refused = () =>
      Result.err(
        new CanonicalSessionError({
          message: "The projection moved on.",
          reason: "refused",
          ground: { condition: "stale" },
        }),
      );
    editor.session.beginVerifiedOperationStaging = (...args) => ({
      ...staging(...args),
      prepare: refused,
      prepareOps: refused,
    });
    const result = editor.apply([
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "alpha",
        replace: "gamma",
      },
    ]);
    expect(result.skipped).toEqual([
      {
        id: "replace",
        reason: "staleRange",
        message: "The projection moved on.",
        canonicalRefusal: { condition: "stale" },
      },
    ]);
  });
});
