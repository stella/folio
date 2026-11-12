// A scan budget is not a subprocess budget.
import { setDefaultTimeout, test } from "bun:test";

import { scanTestTimeout } from "../../scan-timeout";
import { runChild } from "../../subprocess";

setDefaultTimeout(scanTestTimeout(5_000));

test("child", () => runChild(["git", "ls-files"]));
