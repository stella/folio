import { EntityDecoder as XmlEntityDecoder, ENTITY_ACTION } from "@nodable/entities";

// @nodable/entities 3.0 exports a named decoder at runtime, while its declarations
// expose only a default export and omit setXmlVersion, used by fast-xml-parser.
// Remove this augmentation when the published declarations include both.
declare module "@nodable/entities" {
  export class EntityDecoder {
    constructor(options?: EntityDecoderOptions);
    setExternalEntities(entities: Record<string, string>): void;
    addInputEntities(entities: Record<string, string>): void;
    reset(): void;
    decode(text: string): string;
    setXmlVersion(version: string | number): void;
  }
}

// Match fast-xml-parser's existing processEntities expansion-length default.
const MAX_XML_ENTITY_EXPANDED_LENGTH = 100_000;

// Use the typed decoder API: XML names and numeric references, without HTML names.
export const createXmlEntityDecoder = () =>
  new XmlEntityDecoder({
    numericAllowed: true,
    // DOCX package XML cannot declare input entities.
    onInputEntity: () => ENTITY_ACTION.THROW,
    limit: { maxExpandedLength: MAX_XML_ENTITY_EXPANDED_LENGTH, applyLimitsTo: "all" },
  });
