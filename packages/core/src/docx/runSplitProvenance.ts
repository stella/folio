import { TaggedError } from "better-result";
import type { BlockContent, ParagraphContent, PreservedAttribute, Run } from "../types/document";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import {
  getParagraphPropertySourceToken,
  visitDocumentStoryParagraphs,
} from "./paragraphPropertySource";

const OWNER_ATTRIBUTE = "splitRunOwner";
const OWNER_NAMESPACE = OOXML_NAMESPACES.folio.uri;
const OWNER_VALUE = /^1:[^:\s]+:(?<generation>\d+):\d+$/u;

export class InvalidRunSplitProvenanceError extends TaggedError("InvalidRunSplitProvenanceError")<{
  message: string;
}> {}

/** Only our resolved namespace can prove that saved pieces share one source run. */
export const readRunSplitOwner = (
  attributes: readonly PreservedAttribute[] | undefined,
): string | undefined => {
  const owners =
    attributes?.filter(
      ({ namespace, name }) => namespace === OWNER_NAMESPACE && name === OWNER_ATTRIBUTE,
    ) ?? [];
  if (owners.length === 0) return undefined;
  const value = owners.at(0)?.value;
  const match = value === undefined ? null : OWNER_VALUE.exec(value);
  if (
    owners.length !== 1 ||
    match === null ||
    !Number.isSafeInteger(Number(match.groups?.["generation"])) ||
    !Number.isSafeInteger(Number(value?.slice((value?.lastIndexOf(":") ?? -1) + 1)))
  ) {
    throw new InvalidRunSplitProvenanceError({
      message: "Invalid or duplicate saved run-split ownership",
    });
  }
  return value;
};

export const withoutRunSplitOwner = (attributes: readonly PreservedAttribute[] | undefined) =>
  attributes?.filter(
    ({ namespace, name }) => namespace !== OWNER_NAMESPACE || name !== OWNER_ATTRIBUTE,
  );

const runsIn = (content: readonly ParagraphContent[], result: Run[] = []): Run[] => {
  for (const item of content) {
    switch (item.type) {
      case "run":
        result.push(item);
        break;
      case "hyperlink":
        runsIn(item.children, result);
        break;
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      case "inlineWrapper":
      case "inlineSdt":
      case "simpleField":
        runsIn(item.content, result);
        break;
      case "complexField":
        runsIn(item.fieldResult, result);
        break;
    }
  }
  return result;
};

/** Reserve generations across the whole conversion, including nested stories. */
export const runSplitProjectionContext = (blocks: BlockContent[]) => {
  let generation = 0;
  let source: string | undefined;
  visitDocumentStoryParagraphs(blocks, (paragraph) => {
    source ??= getParagraphPropertySourceToken(paragraph);
    for (const run of runsIn(paragraph.content)) {
      const proof = readRunSplitOwner(run.preservedAttributes);
      if (proof !== undefined) {
        const match = OWNER_VALUE.exec(proof);
        generation = Math.max(generation, Number(match?.groups?.["generation"]));
      }
    }
  });
  if (!Number.isSafeInteger(generation + 1)) {
    throw new InvalidRunSplitProvenanceError({
      message: "Saved run-split ownership generation exhausted",
    });
  }
  return { generation: generation + 1, source };
};

export const withRunSplitOwner = (
  attributes: readonly PreservedAttribute[] | undefined,
  proof: string,
): PreservedAttribute[] => [
  ...(withoutRunSplitOwner(attributes) ?? []),
  { namespace: OWNER_NAMESPACE, name: OWNER_ATTRIBUTE, value: proof },
];

type PreserveRunSplitOwnersOptions = {
  content: readonly ParagraphContent[];
  owners: WeakMap<Run, number>;
  proofs: ReadonlyMap<number, string>;
};

/** Persist proof only while one projected owner is split into physical run records. */
export const preserveRunSplitOwners = ({
  content,
  owners,
  proofs,
}: PreserveRunSplitOwnersOptions): void => {
  const groups = new Map<number, Run[]>();
  for (const run of runsIn(content)) {
    const owner = owners.get(run);
    if (owner === undefined) continue;
    const group = groups.get(owner);
    if (group === undefined) groups.set(owner, [run]);
    else group.push(run);
  }
  for (const [owner, pieces] of groups) {
    const proof = proofs.get(owner);
    if (pieces.length < 2 || proof === undefined) continue;
    for (const run of pieces)
      run.preservedAttributes = withRunSplitOwner(run.preservedAttributes, proof);
  }
};
