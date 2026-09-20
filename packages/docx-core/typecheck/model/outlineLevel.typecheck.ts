/**
 * Compile-time proof that `w:outlineLvl` 9 is not a tenth heading level.
 *
 * Nine is the body-text sentinel. The union gives it its own arm, so no
 * heading arm anywhere can carry it. `outlineLevel.property.test.ts` states
 * the runtime half over the whole stated range; the type-level half belongs
 * here, because a package's `typecheck` runs over `tsconfig.build.json`,
 * which excludes `*.test.ts`.
 */

import type { OutlineLevel } from "../../src/model/outlineLevel";

// @ts-expect-error 9 is the body-text sentinel, not a tenth heading level
const HEADING_NINE: OutlineLevel = { kind: "heading", level: 9 };

const HEADING_EIGHT: OutlineLevel = { kind: "heading", level: 8 };

export type OutlineLevelProof = [typeof HEADING_NINE, typeof HEADING_EIGHT];
