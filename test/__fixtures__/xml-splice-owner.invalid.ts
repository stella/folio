// Fixture for `folio-xml-splice/no-hand-rolled-splice`. The rule must flag a
// region cut out of a part, however the concatenation is spelled, and must
// report a chain once rather than once per `+`.

export const viaConcatenation = (xml: string, start: number, end: number, newXml: string): string =>
  xml.slice(0, start) + newXml + xml.slice(end);

export const viaTemplate = (xml: string, start: number, end: number, newXml: string): string =>
  `${xml.slice(0, start)}${newXml}${xml.slice(end)}`;

export const viaLongerChain = (
  xml: string,
  start: number,
  end: number,
  before: string,
  after: string,
): string => xml.slice(0, start) + before + after + xml.slice(end) + "</w:body>";
