import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createDocx } from "../../docx/rezip";
import { revisedFinalParagraphMarks } from "../../compare/verification";
import { expectParagraphAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { schema } from "../schema";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  acceptAllSuggestions,
  acceptSuggestion,
  getSuggestions,
  rejectAIEditRevision,
  rejectAllChanges,
  rejectAllSuggestions,
  rejectSuggestion,
} from "./comments";

const DATE = "2026-09-08T00:00:00.000Z";
const TAIL_REVISION_ID = 701;
const TAIL_SUGGESTION_ID = "insert-tail";
const CARRIER_TEXT = "Existing paragraph.";
const INSERTED_TEXT = "Proposed final paragraph.";
const FIRST_ADJACENT_ID = 801;
const SECOND_ADJACENT_ID = 802;
const FIRST_ADJACENT_SUGGESTION = "first-adjacent-tail";
const SECOND_ADJACENT_SUGGESTION = "second-adjacent-tail";
const FIRST_ADJACENT_TEXT = "First adjacent proposal.";
const SECOND_ADJACENT_TEXT = "Second adjacent proposal.";

type ContainerKind = "body" | "table cell";
type SuggestionAcceptance = "targeted" | "all";
type TrackedResolution = "targeted accept" | "accept all" | "targeted reject" | "reject all";

const paragraph = (
  text: string,
  attrs: Record<string, unknown> = {},
  marks: readonly Mark[] = [],
): PMNode =>
  schema.node("paragraph", attrs, text.length > 0 ? schema.text(text, marks) : undefined);

const tableCell = (paragraphs: PMNode[]): PMNode => schema.node("tableCell", null, paragraphs);

const table = (cell: PMNode): PMNode =>
  schema.node("table", null, [schema.node("tableRow", null, [cell])]);

const alignedParagraph = (
  text: string,
  alignment: "center" | "right",
  attrs: Record<string, unknown> = {},
  marks: readonly Mark[] = [],
): PMNode =>
  paragraph(
    text,
    {
      alignment,
      _originalFormatting: { alignment },
      ...attrs,
    },
    marks,
  );

const makeSuggestedTailState = (containerKind: ContainerKind): EditorState => {
  const insertion = schema.marks["insertion"];
  if (!insertion) {
    return panic("expected the insertion mark type");
  }
  const marker = {
    revisionId: TAIL_REVISION_ID,
    author: "Assistant",
    date: DATE,
    provenance: "suggested" as const,
    suggestionId: TAIL_SUGGESTION_ID,
  };
  const carrier = alignedParagraph(CARRIER_TEXT, "center");
  const inserted = alignedParagraph(INSERTED_TEXT, "right", { _suggestedInsert: marker }, [
    insertion.create(marker),
  ]);
  const blocks =
    containerKind === "body"
      ? [carrier, inserted]
      : [table(tableCell([carrier, inserted])), paragraph("")];
  return EditorState.create({ schema, doc: schema.node("doc", null, blocks) });
};

const suggestedParagraph = ({
  text,
  alignment,
  revisionId,
  suggestionId,
}: {
  text: string;
  alignment: "center" | "right";
  revisionId: number;
  suggestionId: string;
}): PMNode => {
  const insertion = schema.marks["insertion"];
  if (!insertion) {
    return panic("expected the insertion mark type");
  }
  const marker = {
    revisionId,
    author: "Assistant",
    date: DATE,
    provenance: "suggested" as const,
    suggestionId,
  };
  return alignedParagraph(text, alignment, { _suggestedInsert: marker }, [
    insertion.create(marker),
  ]);
};

const makeAdjacentSuggestedTailState = (containerKind: ContainerKind): EditorState => {
  const paragraphs = [
    alignedParagraph(CARRIER_TEXT, "center"),
    suggestedParagraph({
      text: FIRST_ADJACENT_TEXT,
      alignment: "right",
      revisionId: FIRST_ADJACENT_ID,
      suggestionId: FIRST_ADJACENT_SUGGESTION,
    }),
    suggestedParagraph({
      text: SECOND_ADJACENT_TEXT,
      alignment: "center",
      revisionId: SECOND_ADJACENT_ID,
      suggestionId: SECOND_ADJACENT_SUGGESTION,
    }),
  ];
  const blocks =
    containerKind === "body" ? paragraphs : [table(tableCell(paragraphs)), paragraph("")];
  return EditorState.create({ schema, doc: schema.node("doc", null, blocks) });
};

const dispatcher = (state: EditorState) => {
  const view = {
    state,
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

const acceptedSuggestedTail = (containerKind: ContainerKind, acceptance: SuggestionAcceptance) => {
  const view = dispatcher(makeSuggestedTailState(containerKind));
  const command =
    acceptance === "targeted"
      ? acceptSuggestion(TAIL_SUGGESTION_ID, { author: "Reviewer", date: DATE })
      : acceptAllSuggestions({ author: "Reviewer", date: DATE });
  expect(command(view.state, view.dispatch)).toBe(true);
  return view;
};

const paragraphsInChangedContainer = (
  state: EditorState,
  containerKind: ContainerKind,
): PMNode[] => {
  if (containerKind === "body") {
    return Array.from({ length: state.doc.childCount }, (_, index) => state.doc.child(index));
  }
  const cell = state.doc.firstChild?.firstChild?.firstChild;
  if (!cell) {
    return panic("expected the changed table cell");
  }
  return Array.from({ length: cell.childCount }, (_, index) => cell.child(index));
};

const documentBuffer = async (state: EditorState): Promise<Uint8Array> =>
  await createDocx(fromProseDoc(state.doc));

const documentXml = async (state: EditorState): Promise<string> => {
  const buffer = await documentBuffer(state);
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    return panic("expected word/document.xml");
  }
  return await part.async("string");
};

const reopenedState = async (state: EditorState): Promise<EditorState> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await documentBuffer(state));
  return EditorState.create({ schema, doc: toProseDoc(reviewer.toDocument()) });
};

const paragraphXmlContaining = (xml: string, text: string): string => {
  const paragraphs = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gu) ?? [];
  const found = paragraphs.find((candidate) => candidate.includes(text));
  return found ?? panic(`expected a paragraph containing ${text}`);
};

const paragraphProperties = (paragraphXml: string): string =>
  paragraphXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/u)?.at(1) ?? "";

const paragraphPropertyChangeParts = (
  paragraphXml: string,
): { current: string; previous: string } => {
  const changeStart = paragraphXml.indexOf("<w:pPrChange ");
  const propertiesStart = paragraphXml.indexOf("<w:pPr>");
  const previous = paragraphXml
    .match(/<w:pPrChange\b[^>]*><w:pPr>([\s\S]*?)<\/w:pPr><\/w:pPrChange>/u)
    ?.at(1);
  if (propertiesStart < 0 || changeStart < 0 || previous === undefined) {
    return panic("expected a paragraph property change");
  }
  return {
    current: paragraphXml.slice(propertiesStart + "<w:pPr>".length, changeStart),
    previous,
  };
};

const ACCEPTANCE_CASES = (["body", "table cell"] as const).flatMap((containerKind) =>
  (["targeted", "all"] as const).map((acceptance) => ({ acceptance, containerKind })),
);

const TRACKED_RESOLUTIONS = [
  "targeted accept",
  "accept all",
  "targeted reject",
  "reject all",
] as const satisfies readonly TrackedResolution[];

type AdjacentSuggestionDecision = "accept" | "reject";
type AdjacentSuggestionId = typeof FIRST_ADJACENT_SUGGESTION | typeof SECOND_ADJACENT_SUGGESTION;

const ADJACENT_DECISION_CASES = (["body", "table cell"] as const).flatMap((containerKind) =>
  (["accept", "reject"] as const).flatMap((firstDecision) =>
    (["accept", "reject"] as const).map((secondDecision) => ({
      containerKind,
      firstDecision,
      secondDecision,
      label: `${firstDecision} first and ${secondDecision} second in the ${containerKind}`,
    })),
  ),
);

const applySuggestionDecision = (
  view: ReturnType<typeof dispatcher>,
  suggestionId: string,
  decision: AdjacentSuggestionDecision,
): boolean =>
  decision === "accept"
    ? acceptSuggestion(suggestionId, { author: "Reviewer", date: DATE })(view.state, view.dispatch)
    : rejectSuggestion(suggestionId)(view.state, view.dispatch);

const resolveAdjacentSuggestions = (
  containerKind: ContainerKind,
  order: readonly AdjacentSuggestionId[],
  decisions: Readonly<Record<AdjacentSuggestionId, AdjacentSuggestionDecision>>,
) => {
  const view = dispatcher(makeAdjacentSuggestedTailState(containerKind));
  const deferred: AdjacentSuggestionId[] = [];
  for (const suggestionId of order) {
    const decision = decisions[suggestionId];
    if (!applySuggestionDecision(view, suggestionId, decision)) {
      deferred.push(suggestionId);
    }
  }
  expect(deferred).toEqual(
    order[0] === SECOND_ADJACENT_SUGGESTION && decisions[SECOND_ADJACENT_SUGGESTION] === "accept"
      ? [SECOND_ADJACENT_SUGGESTION]
      : [],
  );
  for (const suggestionId of deferred) {
    const decision = decisions[suggestionId];
    expect(applySuggestionDecision(view, suggestionId, decision)).toBe(true);
  }
  expect(getSuggestions(view.state)).toEqual([]);
  expect(revisedFinalParagraphMarks(fromProseDoc(view.state.doc))).toEqual([]);
  return view;
};

describe("accepted suggested container-final paragraphs", () => {
  test.each(ACCEPTANCE_CASES)(
    "$acceptance suggestion acceptance rotates the $containerKind final mark and its formatting",
    async ({ acceptance, containerKind }) => {
      const view = acceptedSuggestedTail(containerKind, acceptance);
      const [carrier, inserted] = paragraphsInChangedContainer(view.state, containerKind);
      if (!carrier || !inserted) {
        return panic("expected the carrier and inserted paragraphs");
      }

      expect(getSuggestions(view.state)).toEqual([]);
      expect(expectParagraphAttrs(carrier).pPrMark).toEqual({
        kind: "ins",
        info: {
          id: TAIL_REVISION_ID,
          author: "Reviewer",
          date: DATE,
        },
      });
      expect(expectParagraphAttrs(inserted)).toMatchObject({
        alignment: "right",
        _originalFormatting: { alignment: "right" },
        _propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: {
              id: TAIL_REVISION_ID,
              author: "Reviewer",
              date: DATE,
            },
            previousFormatting: { alignment: "center" },
          },
        ],
      });
      expect(inserted.attrs["pPrMark"]).toBeNull();

      const model = fromProseDoc(view.state.doc);
      expect(revisedFinalParagraphMarks(model)).toEqual([]);
      const xml = await documentXml(view.state);
      const carrierXml = paragraphXmlContaining(xml, CARRIER_TEXT);
      const insertedXml = paragraphXmlContaining(xml, INSERTED_TEXT);
      expect(paragraphProperties(carrierXml)).toBe(
        `<w:jc w:val="center"/><w:rPr><w:ins w:id="${String(TAIL_REVISION_ID)}" w:author="Reviewer" w:date="${DATE}"/></w:rPr>`,
      );
      expect(paragraphPropertyChangeParts(insertedXml)).toEqual({
        current: '<w:jc w:val="right"/>',
        previous: '<w:jc w:val="center"/>',
      });
      expect(insertedXml).toContain(`<w:pPrChange w:id="0" w:author="Reviewer" w:date="${DATE}">`);
      expect(insertedXml).toContain(`<w:ins w:id="1" w:author="Reviewer" w:date="${DATE}">`);
      expect(paragraphProperties(insertedXml)).not.toContain("<w:rPr>");

      const reopened = await reopenedState(view.state);
      expect(revisedFinalParagraphMarks(fromProseDoc(reopened.doc))).toEqual([]);
      expect(
        paragraphsInChangedContainer(reopened, containerKind).map((node) => ({
          alignment: expectParagraphAttrs(node).alignment,
          directAlignment: expectParagraphAttrs(node)._originalFormatting?.alignment,
          text: node.textContent,
        })),
      ).toEqual([
        { alignment: "center", directAlignment: "center", text: CARRIER_TEXT },
        { alignment: "right", directAlignment: "right", text: INSERTED_TEXT },
      ]);
    },
  );

  test.each(ACCEPTANCE_CASES)(
    "$acceptance suggestion acceptance keeps $containerKind resolution atomic",
    async ({ acceptance, containerKind }) => {
      for (const resolution of TRACKED_RESOLUTIONS) {
        const view = acceptedSuggestedTail(containerKind, acceptance);
        const command = (() => {
          switch (resolution) {
            case "targeted accept":
              return acceptAIEditRevision(TAIL_REVISION_ID);
            case "accept all":
              return acceptAllChanges();
            case "targeted reject":
              return rejectAIEditRevision(TAIL_REVISION_ID);
            case "reject all":
              return rejectAllChanges();
          }
        })();
        expect(command(view.state, view.dispatch)).toBe(true);

        const resolved = paragraphsInChangedContainer(view.state, containerKind);
        const rejectsInsertion = resolution === "targeted reject" || resolution === "reject all";
        expect(resolved.map((node) => node.textContent)).toEqual(
          rejectsInsertion ? [CARRIER_TEXT] : [CARRIER_TEXT, INSERTED_TEXT],
        );
        expect(
          resolved.map((node) => ({
            alignment: expectParagraphAttrs(node).alignment,
            directAlignment: expectParagraphAttrs(node)._originalFormatting?.alignment,
            pPrMark: node.attrs["pPrMark"] ?? null,
            propertyChanges: node.attrs["_propertyChanges"] ?? null,
          })),
        ).toEqual(
          rejectsInsertion
            ? [
                {
                  alignment: "center",
                  directAlignment: "center",
                  pPrMark: null,
                  propertyChanges: null,
                },
              ]
            : [
                {
                  alignment: "center",
                  directAlignment: "center",
                  pPrMark: null,
                  propertyChanges: null,
                },
                {
                  alignment: "right",
                  directAlignment: "right",
                  pPrMark: null,
                  propertyChanges: null,
                },
              ],
        );

        const model = fromProseDoc(view.state.doc);
        expect(revisedFinalParagraphMarks(model)).toEqual([]);
        const xml = await documentXml(view.state);
        expect(xml).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
        expect(paragraphProperties(paragraphXmlContaining(xml, CARRIER_TEXT))).toBe(
          '<w:jc w:val="center"/>',
        );
        if (rejectsInsertion) {
          expect(xml).not.toContain(INSERTED_TEXT);
        } else {
          expect(paragraphProperties(paragraphXmlContaining(xml, INSERTED_TEXT))).toBe(
            '<w:jc w:val="right"/>',
          );
        }

        const reopened = await reopenedState(view.state);
        expect(reopened.doc.textContent).toBe(
          rejectsInsertion ? CARRIER_TEXT : `${CARRIER_TEXT}${INSERTED_TEXT}`,
        );
        expect(revisedFinalParagraphMarks(fromProseDoc(reopened.doc))).toEqual([]);
      }
    },
  );

  test.each(ACCEPTANCE_CASES.filter(({ acceptance }) => acceptance === "all"))(
    "bulk suggestion resolution keeps adjacent tail paragraphs atomic in the $containerKind",
    async ({ containerKind }) => {
      const accepted = dispatcher(makeAdjacentSuggestedTailState(containerKind));
      expect(
        acceptAllSuggestions({ author: "Reviewer", date: DATE })(accepted.state, accepted.dispatch),
      ).toBe(true);

      const targeted = resolveAdjacentSuggestions(
        containerKind,
        [FIRST_ADJACENT_SUGGESTION, SECOND_ADJACENT_SUGGESTION],
        {
          [FIRST_ADJACENT_SUGGESTION]: "accept",
          [SECOND_ADJACENT_SUGGESTION]: "accept",
        },
      );
      expect(accepted.state.doc.toJSON()).toEqual(targeted.state.doc.toJSON());

      const mixed = dispatcher(makeAdjacentSuggestedTailState(containerKind));
      expect(
        acceptSuggestion(FIRST_ADJACENT_SUGGESTION, { author: "Reviewer", date: DATE })(
          mixed.state,
          mixed.dispatch,
        ),
      ).toBe(true);
      expect(rejectAllSuggestions()(mixed.state, mixed.dispatch)).toBe(true);
      const mixedParagraphs = paragraphsInChangedContainer(mixed.state, containerKind);
      expect(mixedParagraphs.map((node) => node.textContent)).toEqual([
        CARRIER_TEXT,
        FIRST_ADJACENT_TEXT,
      ]);
      expect(
        expectParagraphAttrs(mixedParagraphs[0] ?? panic("expected the carrier")).pPrMark,
      ).toEqual({
        kind: "ins",
        info: {
          id: FIRST_ADJACENT_ID,
          author: "Reviewer",
          date: DATE,
        },
      });
      expect(
        expectParagraphAttrs(mixedParagraphs[1] ?? panic("expected the accepted paragraph"))
          ._propertyChanges,
      ).toMatchObject([
        {
          info: { id: FIRST_ADJACENT_ID, author: "Reviewer", date: DATE },
          previousFormatting: { alignment: "center" },
        },
      ]);
      expect(revisedFinalParagraphMarks(fromProseDoc(mixed.state.doc))).toEqual([]);

      const rejected = dispatcher(makeAdjacentSuggestedTailState(containerKind));
      expect(rejectAllSuggestions()(rejected.state, rejected.dispatch)).toBe(true);
      expect(
        paragraphsInChangedContainer(rejected.state, containerKind).map((node) => node.textContent),
      ).toEqual([CARRIER_TEXT]);
      expect(await documentXml(rejected.state)).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
    },
  );

  test.each(ADJACENT_DECISION_CASES)(
    "adjacent suggestions are order-independent when decisions $label",
    async ({ containerKind, firstDecision, secondDecision }) => {
      const decisions = {
        [FIRST_ADJACENT_SUGGESTION]: firstDecision,
        [SECOND_ADJACENT_SUGGESTION]: secondDecision,
      } as const satisfies Readonly<Record<AdjacentSuggestionId, AdjacentSuggestionDecision>>;
      const forward = resolveAdjacentSuggestions(
        containerKind,
        [FIRST_ADJACENT_SUGGESTION, SECOND_ADJACENT_SUGGESTION],
        decisions,
      );
      const reverse = resolveAdjacentSuggestions(
        containerKind,
        [SECOND_ADJACENT_SUGGESTION, FIRST_ADJACENT_SUGGESTION],
        decisions,
      );

      expect(reverse.state.doc.toJSON()).toEqual(forward.state.doc.toJSON());
      expect(await documentXml(reverse.state)).toBe(await documentXml(forward.state));
      const acceptedTexts = [
        ...(firstDecision === "accept" ? [FIRST_ADJACENT_TEXT] : []),
        ...(secondDecision === "accept" ? [SECOND_ADJACENT_TEXT] : []),
      ];
      expect(forward.state.doc.textContent).toBe([CARRIER_TEXT, ...acceptedTexts].join(""));

      for (const pending of [forward, reverse]) {
        const accepting = dispatcher(EditorState.create({ schema, doc: pending.state.doc }));
        const rejecting = dispatcher(EditorState.create({ schema, doc: pending.state.doc }));
        const hasTrackedInsert = acceptedTexts.length > 0;
        if (hasTrackedInsert) {
          expect(acceptAllChanges()(accepting.state, accepting.dispatch)).toBe(true);
          expect(rejectAllChanges()(rejecting.state, rejecting.dispatch)).toBe(true);
        }
        expect(accepting.state.doc.textContent).toBe([CARRIER_TEXT, ...acceptedTexts].join(""));
        expect(rejecting.state.doc.textContent).toBe(CARRIER_TEXT);
        for (const resolved of [accepting, rejecting]) {
          expect(getSuggestions(resolved.state)).toEqual([]);
          expect(revisedFinalParagraphMarks(fromProseDoc(resolved.state.doc))).toEqual([]);
          expect(await documentXml(resolved.state)).not.toMatch(/<w:(?:ins|del|pPrChange)\b/u);
        }
      }
    },
  );

  test("targeted acceptance and rejection leave another suggestion untouched", () => {
    const insertion = schema.marks["insertion"];
    if (!insertion) {
      return panic("expected the insertion mark type");
    }
    const unrelatedMarker = {
      revisionId: 702,
      author: "Assistant",
      date: DATE,
      provenance: "suggested" as const,
      suggestionId: "unrelated-inline",
    };
    const state = makeSuggestedTailState("body");
    const carrier = state.doc.child(0);
    const withUnrelatedSuggestion = carrier.type.create(carrier.attrs, [
      schema.text(CARRIER_TEXT),
      schema.text(" Pending addition.", [insertion.create(unrelatedMarker)]),
    ]);
    const view = dispatcher(
      EditorState.create({
        schema,
        doc: state.doc.type.create(state.doc.attrs, [withUnrelatedSuggestion, state.doc.child(1)]),
      }),
    );

    expect(
      acceptSuggestion(TAIL_SUGGESTION_ID, { author: "Reviewer", date: DATE })(
        view.state,
        view.dispatch,
      ),
    ).toBe(true);
    expect(getSuggestions(view.state).map(({ suggestionId }) => suggestionId)).toEqual([
      "unrelated-inline",
    ]);
    expect(rejectAIEditRevision(TAIL_REVISION_ID)(view.state, view.dispatch)).toBe(true);
    expect(view.state.doc.textContent).toBe(`${CARRIER_TEXT} Pending addition.`);
    expect(getSuggestions(view.state).map(({ suggestionId }) => suggestionId)).toEqual([
      "unrelated-inline",
    ]);
  });
});
