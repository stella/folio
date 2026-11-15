/**
 * Long documents the browser measurements generate instead of checking in:
 * the same seed always gives the same markdown, so every run and every engine
 * measures the same document.
 */

const WORDS =
  "the supplier shall deliver goods services under this agreement subject to terms conditions warranty liability indemnity confidential information party notice period".split(
    " ",
  );

/** A contract of numbered articles, clauses, short lists and fee tables. */
export const contract = (articles: number): string => {
  let seed = 7;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31;
    return seed / 2 ** 31;
  };
  const sentence = (length: number) =>
    Array.from(
      { length },
      () => WORDS[Math.floor(random() * WORDS.length)] ?? "the",
    ).join(" ");
  const lines = ["# Services Agreement", ""];
  for (let article = 1; article <= articles; article += 1) {
    lines.push(`## ${article}. ${sentence(3)}`, "");
    for (let clause = 1; clause <= 8; clause += 1) {
      lines.push(
        `${article}.${clause} **${sentence(2)}** ${sentence(40)}. ${sentence(25)}.`,
        "",
      );
      if (clause % 3 === 0) {
        for (let item = 1; item <= 4; item += 1)
          lines.push(`${item}. ${sentence(18)};`);
        lines.push("");
      }
    }
    if (article % 10 === 0) {
      lines.push("| Item | Fee | Term |", "|---|---|---|");
      for (let row = 0; row < 12; row += 1)
        lines.push(`| ${sentence(3)} | ${row * 100} | ${sentence(2)} |`);
      lines.push("");
    }
  }
  return lines.join("\n");
};

/** One numbered list of short items. */
export const numberedList = (items: number): string =>
  Array.from({ length: items }, (_, index) => `1. Item ${index + 1}`).join(
    "\n",
  );

/** Plain paragraphs only. */
export const plainParagraphs = (paragraphs: number): string =>
  Array.from(
    { length: paragraphs },
    (_, index) => `Paragraph ${index + 1} holds a sentence of plain text.`,
  ).join("\n\n");

/** Many fee tables of short cells, each under a heading. */
export const tables = (count: number, rows: number): string =>
  Array.from({ length: count }, (_, table) =>
    [
      `## Schedule ${table + 1}`,
      "",
      "| Item | Quantity | Fee | Term |",
      "|---|---|---|---|",
      ...Array.from(
        { length: rows },
        (_, row) =>
          `| Item ${table + 1}.${row + 1} | ${row + 1} | ${(row + 1) * 100} | ${row % 12} months |`,
      ),
    ].join("\n"),
  ).join("\n\n");
