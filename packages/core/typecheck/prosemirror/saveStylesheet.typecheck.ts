import type { FromProseDocOptions } from "../../src/prosemirror/conversion/fromProseDoc";

const packageSource = {
  stylesheetSource: { type: "package" },
} as const satisfies FromProseDocOptions;
const suppliedSource = {
  stylesheetSource: { type: "supplied", styles: { styles: [] } },
} as const satisfies FromProseDocOptions;
// @ts-expect-error Saving requires an explicit stylesheet authority.
const missingSource = {} satisfies FromProseDocOptions;
// @ts-expect-error A supplied stylesheet authority must include its stylesheet.
const missingStyles = { stylesheetSource: { type: "supplied" } } satisfies FromProseDocOptions;
void packageSource;
void suppliedSource;
void missingSource;
void missingStyles;
