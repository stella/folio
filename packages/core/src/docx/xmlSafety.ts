import { XMLValidator } from "fast-xml-parser";

const DOCTYPE_PATTERN = /<!DOCTYPE(?:\s|>)/iu;

export type DocxXmlSafetyIssue = "doctype-forbidden" | "not-well-formed";

export const getDocxXmlSafetyIssue = (xml: string): DocxXmlSafetyIssue | null => {
  if (DOCTYPE_PATTERN.test(xml)) {
    return "doctype-forbidden";
  }
  return XMLValidator.validate(xml) === true ? null : "not-well-formed";
};
