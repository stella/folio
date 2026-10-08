import type { ParagraphAttrsPatch } from "../../../src/prosemirror/schema/nodes";
import { listLevelAttrPatch } from "../../../src/prosemirror/styles/resolvedStyleAttrs";

// Patch producers explicitly pass undefined optional attrs. Derive the helper
// input from the schema so exactOptionalPropertyTypes cannot drift between them.
declare const patch: ParagraphAttrsPatch;
const INPUT: Parameters<typeof listLevelAttrPatch>[0] = patch;

export type ListLevelAttrsProof = typeof INPUT;
