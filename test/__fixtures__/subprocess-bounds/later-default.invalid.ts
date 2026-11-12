// A later default replaces the subprocess budget: a bare default and its test.
import { setDefaultTimeout, test } from "bun:test";

import { runChild, subprocessTestTimeout } from "../../subprocess";

setDefaultTimeout(subprocessTestTimeout());
setDefaultTimeout(5_000);

test("child", () => runChild(["bun", "--version"]));
