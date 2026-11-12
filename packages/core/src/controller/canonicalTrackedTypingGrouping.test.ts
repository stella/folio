/**
 * Consecutive tracked typing and deletion by one author read as one change:
 * a typed word saves as one insertion, a run of Backspace or Delete as one
 * deletion, in both engines, whatever the earlier change's date. A paste
 * stays a change of its own, and typing right after it continues it. Another
 * author's change is never continued. Undo takes the edits back exactly, and
 * rejecting gives the source back with its runs whole.
 */
import { describe, expect, test } from "bun:test";
import type { Node as PMNode, Slice } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";

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
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { buildPlainTextSlice } from "../prosemirror/commands/pastePlainText";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, body: string) =>
  `<w:p w14:paraId="${id}">${body}</w:p>`;
const inserted = (id: number, author: string, text: string) =>
  `<w:ins w:id="${id}" w:author="${author}" w:date="2024-01-01T00:00:00Z">${run(text)}</w:ins>`;
const deleted = (id: number, author: string, text: string) =>
  `<w:del w:id="${id}" w:author="${author}" w:date="2024-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

const packageOf = async (body: string) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

type Piece = { kind: "text" | "ins" | "del"; author?: string; text: string };

/**
 * The saved main story's review markup, per paragraph: its mark's review
 * kind and its content as plain text and as each insertion and deletion
 * element on its own, with its author, so that two consecutive changes
 * never read as one. Plain text reads as one piece between them.
 */
const reviewMarkup = async (bytes: Uint8Array) => {
  const xml =
    (await (
      await ZipArchive.read(bytes.slice().buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  const out: { mark: "ins" | "del" | null; content: Piece[] }[] = [];
  const tokens = xml.matchAll(
    /<w:p[ >]|<w:pPr>[\s\S]*?<\/w:pPr>|<(\/?)w:(ins|del)\b([^>]*?)(\/?)>|<w:(t|delText)\b[^>]*>([^<]*)<\/w:\5>/gu,
  );
  let open: Piece | undefined;
  for (const token of tokens) {
    const [whole, closing, revision, attributes, , leaf, text] = token;
    const content = out.at(-1)?.content;
    if (whole.startsWith("<w:p ") || whole === "<w:p>")
      out.push({ mark: null, content: [] });
    else if (whole.startsWith("<w:pPr>")) {
      const mark = /<w:rPr>[\s\S]*<w:(ins|del)\b/u.exec(whole)?.[1];
      const current = out.at(-1);
      if (current !== undefined && (mark === "ins" || mark === "del"))
        current.mark = mark;
    } else if (revision === "ins" || revision === "del") {
      if (closing === "/") open = undefined;
      else {
        open = {
          kind: revision,
          author: /w:author="([^"]*)"/u.exec(attributes ?? "")?.[1] ?? "",
          text: "",
        };
        content?.push(open);
      }
    } else if (leaf !== undefined) {
      const last = content?.at(-1);
      if (open !== undefined) open.text += text ?? "";
      else if (last?.kind === "text") last.text += text ?? "";
      else content?.push({ kind: "text", text: text ?? "" });
    }
  }
  return out;
};

/** How many runs the saved main story holds. */
const runCount = async (bytes: Uint8Array) =>
  [
    ...(
      (await (
        await ZipArchive.read(bytes.slice().buffer)
      )
        .entry("word/document.xml")
        ?.text()) ?? ""
    ).matchAll(/<w:r[ >]/gu),
  ].length;

/** What the saved file reads as pending, accepted and rejected. */
const readings = async (bytes: Uint8Array) => {
  const read = async (decision?: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
    if (decision === "accept") reviewer.acceptAll();
    if (decision === "reject") reviewer.rejectAll();
    return reviewer.snapshot().blocks.map(({ text }) => text);
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

type Step =
  | { caret: readonly [string, number] }
  | { select: readonly [string, number, number] }
  | { type: string }
  | { paste: string }
  | { key: "Backspace" | "Delete"; times: number };

type Driver = {
  state: EditorState;
  dispatch: (tr: EditorState["tr"]) => void;
  typeText: (text: string) => void;
  paste: (slice: Slice) => void;
  pressKey: (binding: string) => unknown;
};

const drive = (driver: Driver, steps: readonly Step[]) => {
  for (const step of steps) {
    if ("caret" in step) {
      const at = positionOf(driver.state.doc, ...step.caret);
      driver.dispatch(
        driver.state.tr.setSelection(
          TextSelection.create(driver.state.doc, at),
        ),
      );
    } else if ("select" in step) {
      const [text, offset, length] = step.select;
      const at = positionOf(driver.state.doc, text, offset);
      driver.dispatch(
        driver.state.tr.setSelection(
          TextSelection.create(driver.state.doc, at, at + length),
        ),
      );
    } else if ("type" in step) driver.typeText(step.type);
    else if ("paste" in step)
      driver.paste(buildPlainTextSlice(step.paste, driver.state.schema));
    else
      for (let count = 0; count < step.times; count += 1)
        driver.pressKey(step.key);
  }
};

const ALPHA = paragraph("11111111", run("Alpha Beta Gamma"));

type Case = {
  id: string;
  body: string;
  steps: readonly Step[];
  /** The saved review elements, as both engines write them. */
  expected: readonly Piece[];
};

const ME = "Conformance";
const CASES: readonly Case[] = [
  {
    id: "a word typed inside a word",
    body: ALPHA,
    steps: [{ caret: ["Beta", 2] }, { type: "XYZ" }],
    expected: [
      { kind: "text", text: "Alpha Be" },
      { kind: "ins", author: ME, text: "XYZ" },
      { kind: "text", text: "ta Gamma" },
    ],
  },
  {
    id: "a word typed at the paragraph end",
    body: ALPHA,
    steps: [{ caret: ["Gamma", 5] }, { type: " Delta" }],
    expected: [
      { kind: "text", text: "Alpha Beta Gamma" },
      { kind: "ins", author: ME, text: " Delta" },
    ],
  },
  {
    id: "typing inside and before a word typed earlier",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 4] },
      { type: "XYZ" },
      { caret: ["XYZ", 1] },
      { type: "uv" },
      { caret: ["XuvYZ", 0] },
      { type: "W" },
    ],
    expected: [
      { kind: "text", text: "Alpha Beta" },
      { kind: "ins", author: ME, text: "WXuvYZ" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "a run of Backspace inside a word",
    body: ALPHA,
    steps: [{ caret: ["Gamma", 4] }, { key: "Backspace", times: 3 }],
    expected: [
      { kind: "text", text: "Alpha Beta G" },
      { kind: "del", author: ME, text: "amm" },
      { kind: "text", text: "a" },
    ],
  },
  {
    id: "a run of Delete from the paragraph start",
    body: ALPHA,
    steps: [{ caret: ["Alpha", 0] }, { key: "Delete", times: 4 }],
    expected: [
      { kind: "del", author: ME, text: "Alph" },
      { kind: "text", text: "a Beta Gamma" },
    ],
  },
  {
    id: "Backspace across a space into the previous word",
    body: ALPHA,
    steps: [{ caret: ["Beta", 4] }, { key: "Backspace", times: 6 }],
    expected: [
      { kind: "text", text: "Alph" },
      { kind: "del", author: ME, text: "a Beta" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "Backspace then Delete at the same place",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 2] },
      { key: "Backspace", times: 2 },
      { key: "Delete", times: 2 },
    ],
    expected: [
      { kind: "text", text: "Alpha " },
      { kind: "del", author: ME, text: "Beta" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "Backspace over a word typed earlier and on into the text",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 4] },
      { type: "XY" },
      { key: "Backspace", times: 4 },
    ],
    expected: [
      { kind: "text", text: "Alpha Be" },
      { kind: "del", author: ME, text: "ta" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "typing over a selection just after the author's own deletion",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 4] },
      { key: "Backspace", times: 2 },
      { select: [" Gamma", 0, 2] },
      { type: "XY" },
    ],
    expected: [
      { kind: "text", text: "Alpha Be" },
      { kind: "del", author: ME, text: "ta G" },
      { kind: "ins", author: ME, text: "XY" },
      { kind: "text", text: "amma" },
    ],
  },
  {
    id: "pasting just after a word typed earlier stays a change of its own",
    body: ALPHA,
    steps: [{ caret: ["Beta", 4] }, { type: "XY" }, { paste: "PQ" }],
    expected: [
      { kind: "text", text: "Alpha Beta" },
      { kind: "ins", author: ME, text: "XY" },
      { kind: "ins", author: ME, text: "PQ" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "pasting just before a word typed earlier stays a change of its own",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 4] },
      { type: "XY" },
      { caret: ["XY", 0] },
      { paste: "PQ" },
    ],
    expected: [
      { kind: "text", text: "Alpha Beta" },
      { kind: "ins", author: ME, text: "PQ" },
      { kind: "ins", author: ME, text: "XY" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "typing just after pasted text continues the pasted insertion",
    body: ALPHA,
    steps: [{ caret: ["Beta", 4] }, { paste: "PQ" }, { type: "XY" }],
    expected: [
      { kind: "text", text: "Alpha Beta" },
      { kind: "ins", author: ME, text: "PQXY" },
      { kind: "text", text: " Gamma" },
    ],
  },
  {
    id: "pasting over a selection just after a word typed earlier",
    body: ALPHA,
    steps: [
      { caret: ["Beta", 4] },
      { type: "XY" },
      { select: [" Gamma", 0, 2] },
      { paste: "PQ" },
    ],
    expected: [
      { kind: "text", text: "Alpha Beta" },
      { kind: "ins", author: ME, text: "XY" },
      { kind: "del", author: ME, text: " G" },
      { kind: "ins", author: ME, text: "PQ" },
      { kind: "text", text: "amma" },
    ],
  },
  {
    id: "typing and deleting beside another author's changes",
    body: paragraph(
      "11111111",
      `${run("Alpha ")}${inserted(5, "Earlier", "Inserted")}${run(" Beta ")}${deleted(6, "Earlier", "Gone")}${run(" Gamma")}`,
    ),
    steps: [
      { caret: ["Inserted", 8] },
      { type: "XY" },
      { caret: [" Gamma", 0] },
      { key: "Delete", times: 2 },
      { caret: [" Beta ", 6] },
      { key: "Backspace", times: 2 },
    ],
    expected: [
      { kind: "text", text: "Alpha " },
      { kind: "ins", author: "Earlier", text: "Inserted" },
      { kind: "ins", author: ME, text: "XY" },
      { kind: "text", text: " Bet" },
      { kind: "del", author: ME, text: "a " },
      { kind: "del", author: "Earlier", text: "Gone" },
      { kind: "del", author: ME, text: " G" },
      { kind: "text", text: "amma" },
    ],
  },
  {
    id: "typing and deleting beside the author's own changes the file came with",
    body: paragraph(
      "11111111",
      `${run("Alpha ")}${inserted(5, ME, "Mine")}${run(" Beta ")}${deleted(6, ME, "Gone")}${run(" Gamma")}`,
    ),
    steps: [
      { caret: ["Mine", 4] },
      { type: "XY" },
      { caret: [" Gamma", 0] },
      { key: "Delete", times: 2 },
    ],
    expected: [
      { kind: "text", text: "Alpha " },
      { kind: "ins", author: ME, text: "MineXY" },
      { kind: "text", text: " Beta " },
      { kind: "del", author: ME, text: "Gone G" },
      { kind: "text", text: "amma" },
    ],
  },
];

const engines = async (body: string, steps: readonly Step[]) => {
  const source = await packageOf(body);
  const parse = () =>
    parseDocx(source, { preloadFonts: false, detectVariables: false });
  const base = await parse();
  const view = new HeadlessEditorView(createHarnessState(base, "suggesting"));
  const harness = createCanonicalEditorHarness(await parse(), "suggesting");
  expect(harness.state.doc.toString()).toBe(view.state.doc.toString());
  drive(view, steps);
  drive(harness, steps);
  return {
    harness,
    expected: (await saveHarnessState(view.state, base)).bytes,
    source: new Uint8Array(source),
  };
};

describe("consecutive tracked edits by one author read as one change", () => {
  for (const { id, body, steps, expected } of CASES)
    test(`${id}: both engines save the same review elements`, async () => {
      const { harness, expected: viaDefault } = await engines(body, steps);
      try {
        expect(harness.refusals).toEqual([]);
        const saved = (await saveCanonicalHarnessDocument(harness.snapshot()))
          .bytes;
        const actual = await readings(saved);
        expect(actual).toEqual(await readings(viaDefault));
        expect(actual.markup).toEqual([{ mark: null, content: [...expected] }]);
      } finally {
        harness.dispose();
      }
    });

  test("undo takes the typed word and the Backspace run back exactly", async () => {
    const { harness, source } = await engines(ALPHA, [
      { caret: ["Beta", 4] },
      { type: "XYZ" },
    ]);
    try {
      const typed = await reviewMarkup(
        (await saveCanonicalHarnessDocument(harness.snapshot())).bytes,
      );
      drive(harness, [{ caret: ["Gamma", 5] }, { key: "Backspace", times: 3 }]);
      const struck = await reviewMarkup(
        (await saveCanonicalHarnessDocument(harness.snapshot())).bytes,
      );
      expect(struck.at(0)?.content.at(-1)).toEqual({
        kind: "del",
        author: ME,
        text: "mma",
      });
      expect(harness.history.undo()).toBe(true);
      expect(
        await reviewMarkup(
          (await saveCanonicalHarnessDocument(harness.snapshot())).bytes,
        ),
      ).toEqual(typed);
      while (harness.history.canUndo()) harness.history.undo();
      const undone = (await saveCanonicalHarnessDocument(harness.snapshot()))
        .bytes;
      expect(await reviewMarkup(undone)).toEqual(await reviewMarkup(source));
      expect(await runCount(undone)).toBe(1);
      while (harness.history.canRedo()) harness.history.redo();
      expect(
        await reviewMarkup(
          (await saveCanonicalHarnessDocument(harness.snapshot())).bytes,
        ),
      ).toEqual(struck);
    } finally {
      harness.dispose();
    }
  });

  test("rejecting a typed word and a Backspace run gives the source back with its run whole", async () => {
    const { harness, source } = await engines(ALPHA, [
      { caret: ["Beta", 2] },
      { type: "XYZ" },
      { caret: ["Gamma", 4] },
      { key: "Backspace", times: 3 },
      { caret: ["Alpha", 0] },
      { key: "Delete", times: 2 },
    ]);
    try {
      expect(harness.resolve("reject")).toBe(true);
      const rejected = (await saveCanonicalHarnessDocument(harness.snapshot()))
        .bytes;
      expect(await reviewMarkup(rejected)).toEqual(await reviewMarkup(source));
      expect(await runCount(rejected)).toBe(1);
    } finally {
      harness.dispose();
    }
  });
});
