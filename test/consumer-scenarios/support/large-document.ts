/** Explicit page boundaries keep document size independent of font measurement. */
import { fromMarkdown } from "@stll/folio-core/markdown";
import { paragraph, run, table } from "@stll/folio-core/server";

import { packFixture } from "./documents.ts";

export const LARGE_DOCUMENT_FIXTURE = "large-100-pages";
const PAGE_COUNT = 100;
const PARAGRAPHS_PER_PAGE = 4;

export const largeDocument = (): Promise<Uint8Array> => {
  const document = fromMarkdown("# Long-flow fixture");
  const content = document.package.document.content;
  content.length = 0;
  for (let page = 0; page < PAGE_COUNT; page += 1) {
    content.push(paragraph(`Page ${page + 1}`, { pageBreakBefore: page > 0 }));
    for (let index = 0; index < PARAGRAPHS_PER_PAGE; index += 1) {
      content.push(
        paragraph([
          run(`Clause ${page + 1}.${index + 1}. `, { bold: true }),
          run("The supplier delivers goods and the buyer provides written notice."),
          run(" Café 日本語 👩🏽‍⚖️."),
        ]),
      );
    }
    if (page % 10 === 0) {
      content.push(
        table({
          header: ["Item", "Value"],
          rows: [
            ["Delivery", "30"],
            ["Notice", "10"],
          ],
        }),
      );
    }
  }
  const boundaries = content.filter(
    (block) => block.type === "paragraph" && block.formatting?.pageBreakBefore,
  );
  if (boundaries.length !== PAGE_COUNT - 1) {
    throw new TypeError("long-flow fixture must declare exactly 100 pages");
  }
  return packFixture(document);
};
