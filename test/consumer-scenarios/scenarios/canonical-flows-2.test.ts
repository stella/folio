/** Shard 2 of the canonical replays of every checked-in flow (support/canonical-replays.ts). */

import { registerCanonicalReplays } from "../support/canonical-replays.ts";

process.env["FOLIO_SCENARIO_ENGINE"] = "canonical";

registerCanonicalReplays(1);
