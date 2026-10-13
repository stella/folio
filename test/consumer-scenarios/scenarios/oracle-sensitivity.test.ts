/** Each targeted run exercises one real oracle against one injected behavioral defect. */
import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
} from "@stll/folio-core/server";

import {
  openReviewer,
  packDocument,
  styleNumberedDocument,
  toArrayBuffer,
} from "../support/documents.ts";
import { GENERATORS } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import { assertReadersAgree, saveAndReopen } from "../support/invariants.ts";
import { ENABLED_RELATIONS, relationCheckCount, startRelations } from "../support/metamorphic.ts";
import {
  ORACLE_MUTATIONS,
  relationForOracle,
  type OracleMutation,
} from "../support/oracle-mutations.ts";
import {
  assertRequestedOutcome,
  capture,
  compareWithModel,
  expectOperation,
  modelOf,
  numberingFactsOf,
  rowsOf,
} from "../support/oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
const prototype = FolioDocxReviewer.prototype;

const injected = process.env["FOLIO_ORACLE_MUTATION"];
if (injected !== undefined && injected !== "" && !Object.hasOwn(ORACLE_MUTATIONS, injected)) {
  throw new TypeError(`Unknown oracle mutation ${injected}`);
}

/** Patch actual reviewer behavior, including the copies and shadows the oracle creates. */
const injectReviewerDefect = (oracle: OracleMutation): (() => void) => {
  const apply = prototype.applyDocumentOperations;
  const undo = prototype.undoDocumentOperations;
  const reject = prototype.rejectAll;
  const getContent = prototype.getContent;
  const addDrift = (reviewer: Reviewer) => {
    const block = getContent.call(reviewer).at(0);
    assert.ok(block, "mutation fixture has an insertion anchor");
    const result = apply.call(
      reviewer,
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "mutation-drift",
            type: "insertAfterBlock",
            blockId: block.id,
            text: "Unrequested paragraph.",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    assert.equal(result.applied.length, 1, "injected drift must apply");
  };
  switch (oracle) {
    case "directTracked":
      // A dynamic receiver is required because shadows and clones use the same prototype.
      prototype.applyDocumentOperations = function (
        this: Reviewer,
        ...args: Parameters<Reviewer["applyDocumentOperations"]>
      ) {
        const [batch, options] = args;
        if (batch.mode !== "tracked-changes") return apply.call(this, batch, options);
        return apply.call(
          this,
          {
            ...batch,
            operations: batch.operations.map((operation) =>
              operation.type === "replaceInBlock"
                ? Object.assign({}, operation, { replace: "Corrupted tracked text" })
                : operation,
            ),
          },
          options,
        );
      };
      break;
    case "rejectAll":
      prototype.rejectAll = prototype.acceptAll;
      break;
    case "undo":
      prototype.undoDocumentOperations = function (
        this: Reviewer,
        ...args: Parameters<Reviewer["undoDocumentOperations"]>
      ) {
        const result = undo.call(this, ...args);
        if (result.status === "undone") addDrift(this);
        return result;
      };
      break;
    case "batchSequential":
      prototype.applyDocumentOperations = function (
        this: Reviewer,
        ...args: Parameters<Reviewer["applyDocumentOperations"]>
      ) {
        const result = apply.call(this, ...args);
        if (args[0].operations.length > 1) addDrift(this);
        return result;
      };
      break;
    case "readerAgreement":
      prototype.getContent = function (
        this: Reviewer,
        ...args: Parameters<Reviewer["getContent"]>
      ) {
        return getContent
          .call(this, ...args)
          .map((block) =>
            Object.assign({}, block, { text: `${block.text} Corrupted reader text.` }),
          );
      };
      break;
    case "saveIdempotent":
    case "readerStability":
    case "requestedOutcome":
    case "saveRoundtrip":
      break;
    default: {
      const exhaustive: never = oracle;
      return exhaustive;
    }
  }
  return () => {
    prototype.applyDocumentOperations = apply;
    prototype.undoDocumentOperations = undo;
    prototype.rejectAll = reject;
    prototype.getContent = getContent;
  };
};

const oracles = Object.keys(ORACLE_MUTATIONS).filter((key): key is OracleMutation =>
  Object.hasOwn(ORACLE_MUTATIONS, key),
);
for (const oracle of oracles) {
  const relation = relationForOracle(oracle);
  test(
    `oracle sensitivity: ${oracle}`,
    { skip: relation !== null && !ENABLED_RELATIONS.has(relation) },
    async () => {
      const fixture = await packDocument(fromMarkdown("First paragraph.\n\nSecond paragraph."));
      const reviewer = await openReviewer(fixture);
      const first = reviewer.getContent().at(0);
      const second = reviewer.getContent().at(1);
      assert.ok(first && second, "sensitivity fixture has two real blocks");
      const batch = {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "first",
            type: "replaceInBlock",
            blockId: first.id,
            find: "First",
            replace: "Updated first",
          },
          {
            id: "second",
            type: "replaceInBlock",
            blockId: second.id,
            find: "Second",
            replace: "Updated second",
          },
        ],
      } as const satisfies Parameters<Reviewer["applyDocumentOperations"]>[0];
      const mutated = injected === oracle;
      const restore = mutated ? injectReviewerDefect(oracle) : () => {};
      try {
        if (oracle === "requestedOutcome") {
          const pre = await capture(reviewer, "direct");
          const result = reviewer.applyDocumentOperations(
            mutated
              ? {
                  ...batch,
                  operations: batch.operations.map((operation) => ({
                    ...operation,
                    replace: "Wrong payload",
                  })),
                }
              : batch,
            { undefinedReferences: "refuse" },
          );
          assert.equal(result.applied.length, 2, "both requested operations must apply");
          await assertRequestedOutcome(
            reviewer,
            pre,
            { applied: batch.operations },
            "oracle sensitivity",
          );
          return;
        }
        if (oracle === "readerAgreement") {
          await assertReadersAgree(fixture, "oracle sensitivity");
          return;
        }
        if (oracle === "saveRoundtrip") {
          if (mutated) {
            const wrong = await packDocument(fromMarkdown("Corrupted saved document."));
            reviewer.toBuffer = () => Promise.resolve(toArrayBuffer(wrong));
          }
          await saveAndReopen(reviewer, "oracle sensitivity");
          return;
        }
        const checksBefore = relation === null ? 0 : relationCheckCount(relation);
        const relations = await startRelations({ fixture, reviewer, mode: "direct", seed: 7 });
        const receipt = reviewer.applyDocumentOperations(batch, { undefinedReferences: "refuse" });
        assert.equal(receipt.applied.length, 2, "both relation operations must apply");
        // The relation itself compares the fault; roundtrip checks must not catch it first.
        const saved = await saveAndReopen(reviewer, "oracle sensitivity", { compare: false });
        if (mutated && oracle === "saveIdempotent") {
          const wrong = await packDocument(fromMarkdown("Corrupted second save."));
          saved.reopened.toBuffer = () => Promise.resolve(toArrayBuffer(wrong));
        }
        if (mutated && oracle === "readerStability") {
          const read = saved.reopened.getContent.bind(saved.reopened);
          saved.reopened.getContent = (...args) =>
            read(...args).map((block) =>
              Object.assign({}, block, { text: `${block.text} Corrupted reopened text.` }),
            );
        }
        await relations.afterStep(reviewer, saved, "oracle sensitivity");
        await relations.finish();
        if (relation !== null && process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"] === "full") {
          assert.ok(
            relationCheckCount(relation) > checksBefore,
            "this probe's selected relation must compare, not skip",
          );
        }
      } finally {
        restore();
      }
    },
  );
}

// The text-only mutation probe cannot expose style-clear numbering drift.
// Exercise every requested numbering source, then restore its stale style tier.
test("oracle sensitivity: style clear across numbering states", async () => {
  const fixture = await styleNumberedDocument();
  const generator = GENERATORS["setBlockParagraphProperties"];
  assert.ok(generator);
  const blocks = (await openReviewer(fixture)).getContent();
  const generatedClearStates = new Set<string>();
  // Pin this seed range independently of the operation matrix's fixture seeds.
  for (let seed = 163_600; seed < 163_856; seed++) {
    const random = createRandom(seed);
    const operation = generator(blocks, random);
    const originalDraws = createRandom(seed);
    // Target, new-list format, existing reference, base property selection.
    for (let draw = 0; draw < 4; draw++) originalDraws.next();
    assert.equal(random.state(), originalDraws.state(), "new dimensions preserve sequence draws");
    const properties = operation?.["properties"];
    if (typeof properties !== "object" || properties === null || !("styleId" in properties))
      continue;
    if (properties.styleId !== null || !("numbering" in properties)) continue;
    const numbering = properties.numbering;
    if (typeof numbering !== "object" || numbering === null || !("kind" in numbering)) continue;
    assert.equal(typeof numbering.kind, "string");
    if (typeof numbering.kind === "string") generatedClearStates.add(numbering.kind);
    if (numbering.kind === "reference") assert.equal("ilvl" in numbering, false);
  }
  assert.deepEqual(generatedClearStates, new Set(["inherit", "levelOnly", "reference"]));
  for (const numberingKind of ["inherit", "levelOnly", "reference"] as const) {
    const reviewer = await openReviewer(fixture);
    const before = rowsOf(reviewer);
    const anchor = before.find((row) => row.styleId === "Heading3");
    assert.ok(anchor?.listReference, "fixture has inherited level-one numbering");
    const numberingByKind = {
      inherit: { kind: "inherit" },
      levelOnly: { kind: "levelOnly", ilvl: 1 },
      reference: paragraphNumberingReference({ numId: anchor.listReference.numId }),
    } as const;
    const operation = {
      id: "clear-style-numbering",
      type: "setBlockParagraphProperties",
      blockId: anchor.id,
      properties: { styleId: null, numbering: numberingByKind[numberingKind] },
    } as const satisfies Parameters<Reviewer["applyDocumentOperations"]>[0]["operations"][number];
    const pre = await capture(reviewer, "direct");
    const receipt = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    );
    assert.equal(receipt.applied.length, 1, "style clear must apply");
    await assertRequestedOutcome(reviewer, pre, { applied: [operation] }, numberingKind);
    const after = rowsOf(reviewer);
    const changed = after.find((row) => row.id === anchor.id);
    assert.ok(changed);
    assert.equal(changed.styleId, undefined);
    assert.deepEqual(
      changed.listReference,
      numberingKind === "reference" ? { numId: anchor.listReference.numId, level: 0 } : undefined,
    );
    const model = modelOf(before);
    model.numberingFacts = await numberingFactsOf(fixture);
    expectOperation(model, operation);
    assert.deepEqual(compareWithModel(model, after), []);
    const staleStyle = after.map((row) =>
      row.id === anchor.id
        ? Object.assign({}, row, {
            listLevel: anchor.listLevel,
            listReference: anchor.listReference,
          })
        : row,
    );
    assert.match(compareWithModel(model, staleStyle).join("\n"), /number|listLevel/u);
  }
});
