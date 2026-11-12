// Four children under the runner's default.
import { beforeAll, test } from "bun:test";

import { runChild } from "../../subprocess";

const check = () => runChild(["bun", "--version"]);
const checkTwice = () => [check(), check()];

beforeAll(() => check());

test("direct", () => Bun.spawnSync(["node", "-e", "0"]));

test("through helpers", () => checkTwice());

test("shell", () => Bun.$`bun --version`.quiet());
