import { panic } from "better-result";

type ResourceReferenceClassificationOptions =
  | {
      kind: "style";
      id: string;
      sourceDefines: (id: string) => boolean;
      destinationDefines: (id: string) => boolean;
    }
  | {
      kind: "numbering";
      id: number;
      sourceDefines: (id: number) => boolean;
      destinationDefines: (id: number) => boolean;
    };

const REFERENCE_DISPOSITIONS = {
  sourceDefined: "sourceDefined",
  destinationCollision: "destinationCollision",
  unknown: "unknown",
} as const;

// A named result keeps declaration emission stable across parallel builds.
type ReferenceDisposition = (typeof REFERENCE_DISPOSITIONS)[keyof typeof REFERENCE_DISPOSITIONS];

type ReferenceSpaceOptions<Id> = {
  id: Id;
  sourceDefines: (id: Id) => boolean;
  destinationDefines: (id: Id) => boolean;
};

const classifyReferenceSpace = <Id>({
  id,
  sourceDefines,
  destinationDefines,
}: ReferenceSpaceOptions<Id>): ReferenceDisposition => {
  if (sourceDefines(id)) return REFERENCE_DISPOSITIONS.sourceDefined;
  return destinationDefines(id)
    ? REFERENCE_DISPOSITIONS.destinationCollision
    : REFERENCE_DISPOSITIONS.unknown;
};

/** A missing source reference may stay dangling, but cannot acquire destination meaning. */
export const classifyResourceReference = (
  reference: ResourceReferenceClassificationOptions,
): ReferenceDisposition => {
  switch (reference.kind) {
    case "style":
      return classifyReferenceSpace(reference);
    case "numbering":
      return classifyReferenceSpace(reference);
    default: {
      const unhandled: never = reference;
      return panic("Unhandled resource reference kind", { reference: unhandled });
    }
  }
};
