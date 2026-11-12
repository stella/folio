import type {
  NumberingDefinitions,
  StyleDefinitions,
  Theme,
} from "../../types/document";
import type { DeepReadonly } from "../../types/documentModel";
import type { ToProseDocOptions } from "./toProseDoc";

/** The package parts a header, footer or note is converted against. */
export type StoryConversionInputs = {
  styles: DeepReadonly<StyleDefinitions> | null | undefined;
  theme: DeepReadonly<Theme> | null | undefined;
  numbering: NumberingDefinitions | null | undefined;
};

/**
 * The options every view of a header, footer or note converts the story with,
 * so the canonical session's projection and the story editors read the same
 * package: its styles, its theme and its lists. A story converted without its
 * numbering keeps its cached markers but loses the indentation its levels give.
 */
export const storyConversionOptions = ({
  styles,
  theme,
  numbering,
}: StoryConversionInputs): ToProseDocOptions => ({
  ...(styles == null ? {} : { styles }),
  ...(theme === undefined ? {} : { theme }),
  ...(numbering == null ? {} : { numbering }),
});
