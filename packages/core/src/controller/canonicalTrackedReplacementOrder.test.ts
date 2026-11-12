/**
 * A tracked replacement reads as a redline: the struck text first, then what
 * replaces it. Typing, pasting or pressing Enter over a selection, and a
 * public replacement, save the same review markup in both engines, and the
 * saved file reads the same once every change is accepted or rejected,
 * including the text a comment covers.
 */
import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createCanonicalEditorHarness,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import {
  createHarnessState,
  HeadlessEditorView,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { buildPlainTextSlice } from "../prosemirror/commands/pastePlainText";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, body: string) =>
  `<w:p w14:paraId="${id}">${body}</w:p>`;
const commented = (word: string) =>
  `<w:commentRangeStart w:id="7"/>${run(word)}<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>`;

const packageOf = async (body: string) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="R" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check this.</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

type Segment = { kind: "text" | "ins" | "del"; text: string };
type Summary = { mark: "ins" | "del" | null; content: Segment[] }[];

/**
 * The review markup of the saved main story, independent of identities: per
 * paragraph, its mark's review kind and its content in order, as plain,
 * inserted and struck text, comment ranges as `[` and `]`. Neighbouring
 * pieces of one kind read as one.
 */
const reviewMarkup = async (bytes: ArrayBuffer): Promise<Summary> => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const summary: Summary = [];
  const tokens = xml.matchAll(
    /<w:p[ >]|<\/w:p>|<w:pPr>[\s\S]*?<\/w:pPr>|<(\/?)w:(ins|del)\b[^>]*?(\/?)>|<w:(t|delText)\b[^>]*>([^<]*)<\/w:\4>|<w:commentRange(Start|End)\b[^>]*\/>/gu,
  );
  let wrapper: "ins" | "del" | null = null;
  const push = (segment: Segment) => {
    const content = summary.at(-1)?.content;
    const previous = content?.at(-1);
    if (previous?.kind === segment.kind) previous.text += segment.text;
    else content?.push(segment);
  };
  for (const token of tokens) {
    const [whole, closing, revision, , leaf, text, range] = token;
    if (whole.startsWith("<w:p ") || whole === "<w:p>")
      summary.push({ mark: null, content: [] });
    else if (whole.startsWith("<w:pPr>")) {
      const mark = /<w:rPr>[\s\S]*<w:(ins|del)\b/u.exec(whole)?.[1];
      const current = summary.at(-1);
      if (current !== undefined && (mark === "ins" || mark === "del"))
        current.mark = mark;
    } else if (revision === "ins" || revision === "del")
      wrapper = closing === "/" ? null : revision;
    else if (leaf !== undefined)
      push({ kind: wrapper ?? "text", text: text ?? "" });
    else if (range !== undefined)
      push({ kind: "text", text: range === "Start" ? "[" : "]" });
  }
  return summary;
};

/** What the saved file reads as pending, accepted and rejected. */
const readings = async (bytes: ArrayBuffer) => {
  const read = async (decision?: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    if (decision === "accept") reviewer.acceptAll();
    if (decision === "reject") reviewer.rejectAll();
    return {
      text: reviewer.snapshot().blocks.map(({ text }) => text),
      comments: reviewer
        .getComments()
        .map(({ id, anchoredText }) => ({ id, anchoredText })),
    };
  };
  return {
    markup: await reviewMarkup(bytes),
    pending: await read(),
    accepted: await read("accept"),
    rejected: await read("reject"),
  };
};

/** The editor position just before the `skip`th character of `needle`. */
const positionOf = (doc: PMNode, needle: string, skip = 0): number => {
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    const index = node.isText ? (node.text ?? "").indexOf(needle) : -1;
    if (index !== -1) found = position + index + skip;
    return true;
  });
  if (found === undefined) throw new TypeError(`No text ${needle}`);
  return found;
};

type EditorEdit = "type" | "paste" | "enter";
type Driver = {
  state: EditorState;
  dispatch: (tr: EditorState["tr"]) => void;
  typeText: (text: string) => void;
  paste: (slice: ReturnType<typeof buildPlainTextSlice>) => void;
  pressKey: (binding: string) => unknown;
};

type EditorCase = {
  id: string;
  body: string;
  from: readonly [string, number];
  to: readonly [string, number];
  edit: EditorEdit;
};

const EDITOR_CASES: readonly EditorCase[] = [
  {
    id: "typing over a word",
    body: paragraph("11111111", run("Alpha Beta Gamma")),
    from: ["Beta", 0],
    to: ["Beta", 4],
    edit: "type",
  },
  {
    id: "typing over a selection across paragraphs",
    body:
      paragraph("11111111", run("Alpha Beta")) +
      paragraph("22222222", run("Gamma Delta")),
    from: ["Beta", 2],
    to: ["Gamma", 3],
    edit: "type",
  },
  {
    id: "pasting two lines over a selection across paragraphs",
    body:
      paragraph("11111111", run("Alpha Beta")) +
      paragraph("22222222", run("Gamma Delta")),
    from: ["Beta", 2],
    to: ["Gamma", 3],
    edit: "paste",
  },
  {
    id: "pasting two lines over a word",
    body: paragraph("11111111", run("Alpha Beta Gamma")),
    from: ["Beta", 0],
    to: ["Beta", 4],
    edit: "paste",
  },
  {
    id: "pressing Enter over a word",
    body: paragraph("11111111", run("Alpha Beta Gamma")),
    from: ["Beta", 0],
    to: ["Beta", 4],
    edit: "enter",
  },
  {
    id: "typing over a selection ending inside a content control",
    body: paragraph(
      "11111111",
      `${run("Alpha ")}<w:sdt><w:sdtPr><w:id w:val="100"/></w:sdtPr><w:sdtContent>${run("Beta Gamma")}</w:sdtContent></w:sdt>`,
    ),
    from: ["Alpha ", 6],
    to: ["Beta", 2],
    edit: "type",
  },
  {
    id: "typing over a selection ending inside a comment",
    body: paragraph(
      "11111111",
      `${run("Notes on ")}${commented("Documentation")}${run(" of Training")}`,
    ),
    from: ["Notes", 2],
    to: ["Documentation", 3],
    edit: "type",
  },
];

const drive = (driver: Driver, { from, to, edit }: EditorCase) => {
  const doc = driver.state.doc;
  driver.dispatch(
    driver.state.tr.setSelection(
      TextSelection.create(
        doc,
        positionOf(doc, ...from),
        positionOf(doc, ...to),
      ),
    ),
  );
  if (edit === "type") driver.typeText("XY");
  if (edit === "paste")
    driver.paste(buildPlainTextSlice("one\ntwo", driver.state.schema));
  if (edit === "enter") driver.pressKey("Enter");
};

const editorReadings = async (editorCase: EditorCase) => {
  const source = await packageOf(editorCase.body);
  const parse = () =>
    parseDocx(source, { preloadFonts: false, detectVariables: false });
  const base = await parse();
  const initial = createHarnessState(base, "suggesting");
  const view = new HeadlessEditorView(initial);
  const harness = createCanonicalEditorHarness(await parse(), "suggesting");
  try {
    // Both engines show the same document, so one selection names one range.
    expect(harness.state.doc.toString()).toBe(initial.doc.toString());
    drive(view, editorCase);
    drive(harness, editorCase);
    expect(harness.refusals).toEqual([]);
    const expected = (await saveHarnessState(view.state, base)).bytes;
    const actual = (await saveCanonicalHarnessDocument(harness.snapshot()))
      .bytes;
    return {
      expected: await readings(expected.slice().buffer),
      actual: await readings(actual.slice().buffer),
    };
  } finally {
    harness.dispose();
  }
};

/** The public replacement of `find` by `replace` in a tracked batch. */
const operation = (find: string, replace: string) =>
  ({
    id: "edit",
    type: "replaceInBlock",
    blockId: "11111111",
    find,
    replace,
  }) as const;

const viaDefaultEngine = async (
  source: ArrayBuffer,
  find: string,
  replace: string,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  const result = reviewer.applyOperations([operation(find, replace)], {
    mode: "tracked-changes",
    undefinedReferences: "refuse",
  });
  expect(result.applied).toHaveLength(1);
  return reviewer.toBuffer();
};

const viaCanonicalEngine = async (
  source: ArrayBuffer,
  find: string,
  replace: string,
) => {
  const document = await parseDocx(source, {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: {
      version: 1,
      mode: "tracked-changes",
      operations: [operation(find, replace)],
    },
    undefinedReferences: "refuse",
  });
  expect(result.applied).toHaveLength(1);
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

const PUBLIC_CASES = [
  {
    id: "a public replacement of a word",
    body: paragraph("11111111", run("Alpha Beta Gamma")),
    find: "Beta",
  },
  {
    id: "a public replacement of exactly the text a comment covers",
    body: paragraph(
      "11111111",
      `${run("Notes on ")}${commented("Documentation")}${run(" of Training")}`,
    ),
    find: "Documentation",
  },
] as const;

describe("a tracked replacement puts the struck text first", () => {
  for (const editorCase of EDITOR_CASES)
    test(`${editorCase.id}: both engines save the same redline`, async () => {
      const { expected, actual } = await editorReadings(editorCase);
      expect(actual).toEqual(expected);
      const content = expected.markup.flatMap(({ content: pieces }) => pieces);
      const struck = content.findIndex(({ kind }) => kind === "del");
      const inserted = content.findIndex(({ kind }) => kind === "ins");
      expect(struck).toBeGreaterThanOrEqual(0);
      if (inserted !== -1) expect(struck).toBeLessThan(inserted);
    });

  test("typing over exactly the text a comment covers keeps the comment on the new text", async () => {
    const body = paragraph(
      "11111111",
      `${run("Notes on ")}${commented("Documentation")}${run(" of Training")}`,
    );
    const source = await packageOf(body);
    const harness = createCanonicalEditorHarness(
      await parseDocx(source, { preloadFonts: false, detectVariables: false }),
      "suggesting",
    );
    try {
      drive(harness, {
        id: "commented word",
        body,
        from: ["Documentation", 0],
        to: ["Documentation", 13],
        edit: "type",
      });
      expect(harness.refusals).toEqual([]);
      const saved = await saveCanonicalHarnessDocument(harness.snapshot());
      const actual = await readings(saved.bytes.slice().buffer);
      // As the public replacement of the same words saves in both engines.
      // The default editor's typing puts the new text after the comment's
      // end instead, so accepting there leaves the comment nothing to cover.
      const replaced = await readings(
        await viaDefaultEngine(source, "Documentation", "XY"),
      );
      expect(actual).toEqual(replaced);
      expect(actual.markup.at(0)?.content).toEqual([
        { kind: "text", text: "Notes on [" },
        { kind: "del", text: "Documentation" },
        { kind: "ins", text: "XY" },
        { kind: "text", text: "] of Training" },
      ]);
      expect(actual.accepted.comments).toEqual([{ id: 7, anchoredText: "XY" }]);
    } finally {
      harness.dispose();
    }
  });

  for (const { id, body, find } of PUBLIC_CASES)
    test(`${id}: both engines save the same redline`, async () => {
      const source = await packageOf(body);
      const expected = await readings(
        await viaDefaultEngine(source, find, "Changed"),
      );
      const actual = await readings(
        await viaCanonicalEngine(source, find, "Changed"),
      );
      expect(actual).toEqual(expected);
      expect(
        expected.markup.at(0)?.content.filter(({ kind }) => kind !== "text"),
      ).toEqual([
        { kind: "del", text: find },
        { kind: "ins", text: "Changed" },
      ]);
    });
});
