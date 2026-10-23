// Bun's test loader transpiles fixtures without typechecking. Compile the total
// source-kind census here so new model members require a fixture disposition.
import {
  PARAGRAPH_SEARCH_SOURCE_CENSUS,
  RUN_SEARCH_SOURCE_CENSUS,
} from "../../../test/findReplacementBoundaryFixtures";

void RUN_SEARCH_SOURCE_CENSUS;
void PARAGRAPH_SEARCH_SOURCE_CENSUS;
