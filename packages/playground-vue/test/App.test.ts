import { expect, test } from "bun:test";
import { panic } from "better-result";
import { compileTemplate, parse } from "vue/compiler-sfc";

const roots = [new URL("../src/", import.meta.url), new URL("../../vue/src/", import.meta.url)];
const components = roots.flatMap((root) =>
  [...new Bun.Glob("**/*.vue").scanSync({ cwd: root.pathname })]
    .sort()
    .map((filename) => new URL(filename, root)),
);

test.each(components)("Vue component %s parses without template errors", async (url) => {
  const source = await Bun.file(url).text();
  expect(parse(source, { filename: url.pathname }).errors).toEqual([]);
});

test("the playground entry compiles before either editor session mounts", async () => {
  const filename = new URL("../src/App.vue", import.meta.url).pathname;
  const source = await Bun.file(filename).text();
  const parsed = parse(source, { filename });
  expect(parsed.errors).toEqual([]);
  const template = parsed.descriptor.template ?? panic("The playground must have a template.");
  const compiled = compileTemplate({ filename, id: "playground", source: template.content });
  expect(compiled.errors).toEqual([]);
});
