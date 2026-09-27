import JSZip from "jszip";

/** Finite, deterministic malformed XML and cross-part package mutations. */
export const PARSER_XML_MUTATIONS = [
  "dropElement",
  "duplicateElement",
  "reorderElements",
  "invalidAttribute",
  "outOfRangeId",
  "deepNesting",
  "hugeText",
  "emptyText",
  "unknownNamespace",
  "bom",
  "encoding",
  "danglingNumberingId",
  "danglingStyleId",
  "danglingCommentId",
  "danglingFootnoteId",
  "cyclicStyles",
  "selfReferencingRelationship",
] as const;

export type ParserXmlMutation = (typeof PARSER_XML_MUTATIONS)[number];

const MAX_INPUT_BYTES = 32 * 1024 * 1024;
const MAX_OUTPUT_BYTES = 40 * 1024 * 1024;
const MAX_PART_BYTES = 8 * 1024 * 1024;
const HUGE_TEXT_CHARS = 256 * 1024;
const DEEP_NESTING_LEVEL = 120;
const XML_DECLARATION = /^\uFEFF?<\?xml\b[^?]*\?>/iu;

const requiredMatch = (xml: string, expression: RegExp, label: string): RegExpExecArray => {
  const match = expression.exec(xml);
  if (!match) throw new Error(`Cannot apply XML mutation: no ${label} found`);
  return match;
};

const replaceFirst = (
  xml: string,
  expression: RegExp,
  replacement: string | ((match: string) => string),
  label: string,
): string => {
  requiredMatch(xml, expression, label);
  return xml.replace(expression, replacement);
};

const mutateXmlText = (xml: string, mutation: ParserXmlMutation, targetPart: string): string => {
  switch (mutation) {
    case "dropElement": {
      const match = requiredMatch(
        xml,
        /<w:p\b[^>]*>[\s\S]*?<\/w:p>|<w:r\b[^>]*>[\s\S]*?<\/w:r>|<([\w.-]+:[\w.-]+|[\w.-]+)\b[^>]*\/>/u,
        "child element",
      );
      return xml.replace(match[0], "");
    }
    case "duplicateElement": {
      const match = requiredMatch(
        xml,
        /<w:p\b[^>]*>[\s\S]*?<\/w:p>|<w:r\b[^>]*>[\s\S]*?<\/w:r>|<([\w.-]+:[\w.-]+|[\w.-]+)\b[^>]*\/>/u,
        "child element",
      );
      return xml.replace(match[0], `${match[0]}${match[0]}`);
    }
    case "reorderElements": {
      const emptyElements = /(<[^!?/][^>]*\/>)\s*(<[^!?/][^>]*\/>)/u;
      const blockElements =
        /(<w:(?:p|tr|tc|tbl|style|abstractNum|num)\b[^>]*>[\s\S]*?<\/w:(?:p|tr|tc|tbl|style|abstractNum|num)>)(\s*)(<w:(?:p|tr|tc|tbl|style|abstractNum|num)\b[^>]*>[\s\S]*?<\/w:(?:p|tr|tc|tbl|style|abstractNum|num)>)/u;
      const empty = emptyElements.exec(xml);
      if (empty) return xml.replace(empty[0], `${empty[2]}${empty[1]}`);
      const block = blockElements.exec(xml);
      if (block) return xml.replace(block[0], `${block[3]}${block[2]}${block[1]}`);
      throw new Error("Cannot apply XML mutation: no reorderable adjacent elements found");
    }
    case "invalidAttribute": {
      const match = requiredMatch(xml, /<[\w.-]+:[\w.-]+\b|<[\w.-]+\b/u, "start tag");
      return (
        xml.slice(0, match.index + match[0].length) +
        ' broken="\u0000"' +
        xml.slice(match.index + match[0].length)
      );
    }
    case "outOfRangeId": {
      const match = requiredMatch(
        xml,
        /\b(?:w:)?(?:id|numId|styleId)\s*=\s*(["'])[^"']*\1/iu,
        "id attribute",
      );
      return xml.replace(match[0], match[0].replace(/(["'])[^"']*\1/u, '"2147483648"'));
    }
    case "deepNesting": {
      const wrapper = "<w:proofErr>";
      const close = "</w:proofErr>";
      if (/<w:body\b[^>]*>[\s\S]*?<\/w:body>/u.test(xml)) {
        return xml
          .replace(
            /<w:body\b[^>]*>/u,
            (opening) => `${opening}${wrapper.repeat(DEEP_NESTING_LEVEL)}`,
          )
          .replace(/<\/w:body>/u, `${close.repeat(DEEP_NESTING_LEVEL)}</w:body>`);
      }
      const root = requiredMatch(xml, /<([\w.-]+:[\w.-]+|[\w.-]+)\b[^>]*>/u, "root element");
      const rootName = root[1];
      const rootClose = new RegExp(`<\\/${rootName}\\s*>`, "u");
      requiredMatch(xml, rootClose, "root closing tag");
      return xml
        .replace(root[0], `${root[0]}${wrapper.repeat(DEEP_NESTING_LEVEL)}`)
        .replace(rootClose, `${close.repeat(DEEP_NESTING_LEVEL)}$&`);
    }
    case "hugeText": {
      const text = "x".repeat(HUGE_TEXT_CHARS);
      if (/<[^!?/][^>]*>[\s\S]*?<\//u.test(xml)) {
        return xml.replace(/(<w:t\b[^>]*>)[\s\S]*?(<\/w:t>)/u, `$1${text}$2`);
      }
      return xml.replace(/<\/[^>]+>/u, `${text}$&`);
    }
    case "emptyText": {
      return replaceFirst(xml, /(<w:t\b[^>]*>)[\s\S]*?(<\/w:t>)/u, "$1$2", "w:t text node");
    }
    case "unknownNamespace": {
      return replaceFirst(
        xml,
        /xmlns:w\s*=\s*(["'])[^"']*\1/u,
        'xmlns:w="urn:folio:unknown-wordprocessingml"',
        "w namespace declaration",
      );
    }
    case "bom":
      if (xml.startsWith("\uFEFF"))
        throw new Error("Cannot apply BOM mutation: target already has a UTF-8 BOM");
      return `\uFEFF${xml}`;
    case "encoding": {
      const declaration = XML_DECLARATION.exec(xml);
      const changed = declaration
        ? declaration[0].replace(/encoding\s*=\s*(["'])[^"']*\1/iu, 'encoding="UTF-16"')
        : '<?xml version="1.0" encoding="UTF-16"?>';
      return declaration ? xml.replace(XML_DECLARATION, changed) : `${changed}${xml}`;
    }
    case "danglingNumberingId":
      return replaceFirst(
        xml,
        /<w:numId\b[^>]*w:val\s*=\s*(["'])\d+\1[^>]*\/?\s*>/u,
        (match) => match.replace(/w:val\s*=\s*(["'])\d+\1/u, 'w:val="2147483647"'),
        "w:numId reference",
      );
    case "danglingStyleId":
      return replaceFirst(
        xml,
        /<w:(?:pStyle|rStyle|tblStyle)\b[^>]*w:val\s*=\s*(["'])[^"']+\1[^>]*\/?\s*>/u,
        (match) => match.replace(/w:val\s*=\s*(["'])[^"']+\1/u, 'w:val="FolioMissingStyle"'),
        "style reference",
      );
    case "danglingCommentId":
      return replaceFirst(
        xml,
        /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*w:id\s*=\s*(["'])-?\d+\1[^>]*\/?\s*>/u,
        (match) => match.replace(/w:id\s*=\s*(["'])-?\d+\1/u, 'w:id="2147483647"'),
        "comment reference",
      );
    case "danglingFootnoteId":
      return replaceFirst(
        xml,
        /<w:footnoteReference\b[^>]*w:id\s*=\s*(["'])-?\d+\1[^>]*\/?\s*>/u,
        (match) => match.replace(/w:id\s*=\s*(["'])-?\d+\1/u, 'w:id="2147483647"'),
        "footnote reference",
      );
    case "cyclicStyles": {
      const styles = /<w:style\b[^>]*w:styleId\s*=\s*(["'])([^"']+)\1[^>]*>[\s\S]*?<\/w:style>/gu;
      const basedOn = /<w:basedOn\b[^>]*w:val\s*=\s*(["'])[^"']+\1[^>]*\/?\s*>/u;
      let changed = false;
      const updated = xml.replace(styles, (style) => {
        if (changed) return style;
        const opening = /<w:style\b[^>]*w:styleId\s*=\s*(["'])([^"']+)\1[^>]*>/u.exec(style);
        const styleId = opening?.[2];
        if (!opening || !styleId) return style;
        const current = basedOn.exec(style);
        if (current?.[0].includes(`w:val="${styleId}"`)) return style;
        changed = true;
        return current
          ? style.replace(
              current[0],
              current[0].replace(/w:val\s*=\s*(["'])[^"']+\1/u, `w:val="${styleId}"`),
            )
          : style.replace(opening[0], `${opening[0]}<w:basedOn w:val="${styleId}"/>`);
      });
      if (!changed)
        throw new Error("Cannot apply XML mutation: no style without a self-reference found");
      return updated;
    }
    case "selfReferencingRelationship": {
      if (targetPart.endsWith(".rels")) {
        const selfTarget =
          targetPart === "_rels/.rels"
            ? targetPart
            : `_rels/${targetPart.split("/_rels/").at(-1) ?? targetPart}`;
        return replaceFirst(
          xml,
          /(<Relationship\b[^>]*\bTarget\s*=\s*)(["'])[^"']*\2/u,
          (match) => {
            const attributeStart = match.slice(0, match.indexOf("Target"));
            const quote = match.includes("Target='") ? "'" : '"';
            return `${attributeStart}Target=${quote}${selfTarget}${quote}`;
          },
          "relationship target",
        );
      }
      throw new Error(`Cannot mutate relationship XML in ${targetPart}; expected a .rels part`);
    }
    default: {
      const exhaustive: never = mutation;
      throw new Error(`Unknown parser XML mutation ${String(exhaustive)}`);
    }
  }
};

/** Apply one bounded deterministic XML mutation to a part in a DOCX ZIP. */
export const mutateXml = async (
  source: Uint8Array,
  mutation: ParserXmlMutation,
  targetPart: string,
): Promise<Uint8Array> => {
  if (source.byteLength > MAX_INPUT_BYTES)
    throw new Error("DOCX mutation input exceeds size limit");
  const zip = await JSZip.loadAsync(source);
  const file = zip.file(targetPart);
  if (!file) throw new Error(`DOCX mutation target part not found: ${targetPart}`);
  const original = await file.async("uint8array");
  if (original.byteLength > MAX_PART_BYTES)
    throw new Error(`DOCX mutation part exceeds size limit: ${targetPart}`);
  const decoded = new TextDecoder("utf-8").decode(original);
  const changedXml = mutateXmlText(decoded, mutation, targetPart);
  if (changedXml === decoded)
    throw new Error(`XML mutation ${mutation} made no change to ${targetPart}`);
  let changed: Uint8Array;
  if (mutation === "encoding") {
    const utf16 = new Uint8Array(2 + changedXml.length * 2);
    utf16[0] = 0xff;
    utf16[1] = 0xfe;
    for (let index = 0; index < changedXml.length; index += 1) {
      const codeUnit = changedXml.charCodeAt(index);
      utf16[2 + index * 2] = codeUnit & 0xff;
      utf16[3 + index * 2] = codeUnit >> 8;
    }
    changed = utf16;
  } else {
    changed = new TextEncoder().encode(changedXml);
  }
  if (changed.byteLength > MAX_PART_BYTES)
    throw new Error("DOCX mutation output part exceeds size limit");
  zip.file(targetPart, changed, { date: new Date("1980-01-01T00:00:00.000Z") });
  const output = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  if (output.byteLength > MAX_OUTPUT_BYTES)
    throw new Error("DOCX mutation output exceeds size limit");
  return output;
};
