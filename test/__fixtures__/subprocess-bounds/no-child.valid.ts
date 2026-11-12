// A file that starts no child is not this rule's concern.
import { setDefaultTimeout, test } from "bun:test";

setDefaultTimeout(30_000);

test("in process", () => {}, 5_000);
