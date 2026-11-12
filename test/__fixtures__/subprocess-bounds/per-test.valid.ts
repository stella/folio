// Under another default, each test that starts a child states its own budget.
import { setDefaultTimeout, test } from "bun:test";

import { propertyTestTimeout } from "../../property-timeout";
import { runChild, subprocessTestTimeout } from "../../subprocess";

setDefaultTimeout(propertyTestTimeout(5_000));

const check = () => runChild(["bun", "--version"]);

test("property", () => {});

test("child", () => check(), subprocessTestTimeout());
