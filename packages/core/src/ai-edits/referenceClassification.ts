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

type ReferenceSpaceOptions<Id> = {
  id: Id;
  sourceDefines: (id: Id) => boolean;
  destinationDefines: (id: Id) => boolean;
};

const classifyReferenceSpace = <Id>({
  id,
  sourceDefines,
  destinationDefines,
}: ReferenceSpaceOptions<Id>) => {
  if (sourceDefines(id)) return "sourceDefined";
  return destinationDefines(id) ? "destinationCollision" : "unknown";
};

/** A missing source reference may stay dangling, but cannot acquire destination meaning. */
export const classifyResourceReference = (reference: ResourceReferenceClassificationOptions) => {
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
