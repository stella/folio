import {
  ENTITY_ACTION,
  EntityDecoder as XmlEntityDecoder,
} from "@nodable/entities";

// Match fast-xml-parser's existing processEntities expansion-length default.
const MAX_XML_ENTITY_EXPANDED_LENGTH = 100_000;

// Use the typed decoder API: XML names and numeric references, without HTML names.
export const createXmlEntityDecoder = (): XmlEntityDecoder =>
  new XmlEntityDecoder({
    numericAllowed: true,
    // DOCX package XML cannot declare input entities.
    onInputEntity: () => ENTITY_ACTION.THROW,
    limit: {
      maxExpandedLength: MAX_XML_ENTITY_EXPANDED_LENGTH,
      applyLimitsTo: "all",
    },
  });
