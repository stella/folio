import type {
  FolioAIDiagnosticBlock,
  FolioAIParagraphBlock,
  FolioAIEditOperation,
  FolioAIResolvedEditOperation,
} from "../../src/ai-edits/types";
import { resolveNewListOperations } from "../../src/ai-edits/newListNumbering";
import { listLevelAttrPatch } from "../../src/prosemirror/styles/resolvedStyleAttrs";

const request = {
  id: "numbering-request",
  type: "setBlockParagraphProperties",
  blockId: "paragraph",
  properties: { numbering: { kind: "newList", format: "numbered" } },
} as const satisfies FolioAIEditOperation;

// @ts-expect-error the allocator must resolve newList before application
const unresolved: FolioAIResolvedEditOperation = request;

const resolved = resolveNewListOperations([request], null).operations;
const applied: readonly FolioAIResolvedEditOperation[] = resolved;

export type ResolvedNumberingProof = typeof applied | typeof unresolved;

// Level changes derive the current list; changing its identity requires a reference request.

const level: Parameters<typeof listLevelAttrPatch>[1] = 1;
// @ts-expect-error a level change cannot request another list identity
const differentList: Parameters<typeof listLevelAttrPatch>[1] = { numId: 2, ilvl: 1 };

export type SameListLevelProof = typeof level | typeof differentList;

// Diagnostic carriers cannot invent a paragraph numbering source.
declare const diagnostic: FolioAIDiagnosticBlock;
// @ts-expect-error an opaque carrier has no authored numbering
const diagnosticSource = diagnostic.statedNumbering;
// @ts-expect-error an opaque carrier has no effective list membership
const diagnosticList = diagnostic.listReference;
declare const paragraph: FolioAIParagraphBlock;
const paragraphSource = paragraph.statedNumbering;

export type BlockNumberingProof =
  | typeof diagnosticSource
  | typeof diagnosticList
  | typeof paragraphSource;
