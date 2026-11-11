# Scratch only: add a timer (globalThis.zzT, ms) around the operations contract check in applyDocumentOp.
# usage: python3 -I instrument.py <tree>
import sys
p=sys.argv[1]+'/packages/docx-core/src/ops/apply.ts'
s=open(p).read()
a='  const valid = validateOpsDocument(document);\n'
assert s.count(a)==1
s=s.replace(a,'  const zzValid = performance.now();\n'+a+'  zzAdd("validateInApply", zzValid);\n')
s+='''
function zzAdd(k: string, t: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + performance.now() - t;
}
'''
open(p,'w').write(s)
print('instrumented', p)
p=sys.argv[1]+'/packages/core/src/docx/listModelNormalization.ts'
s=open(p).read()
base_a='    return content === undefined\n      ? body\n      : { ...body, ...withBodyContent(body, content) };\n'
head_a='  for (const { story, body } of stories)\n    normalized = withReplaced(normalized, story, body, replacedIn(body, story));\n'
if s.count(base_a)==1:
    s=s.replace(base_a,'    if (content === undefined) return body;\n    const zzPlace = performance.now();\n    const zzBody = { ...body, ...withBodyContent(body, content) };\n    zzAdd("placeReplaced", zzPlace);\n    return zzBody;\n')
    print('placeReplaced: base anchor')
elif s.count(head_a)==1:
    s=s.replace(head_a,'  for (const { story, body } of stories) {\n    const zzReplaced = replacedIn(body, story);\n    const zzPlace = performance.now();\n    normalized = withReplaced(normalized, story, body, zzReplaced);\n    zzAdd("placeReplaced", zzPlace);\n  }\n')
    print('placeReplaced: head anchor')
else:
    print('placeReplaced: no anchor')
s+='''
function zzAdd(k: string, t: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + performance.now() - t;
}
function zzCount(k: string, n: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + n;
}
const zzStory = ListNormalizationCache.prototype.story;
ListNormalizationCache.prototype.story = function (
  this: ListNormalizationCache,
  ...args: Parameters<typeof zzStory>
) {
  const t = performance.now();
  const r = zzStory.apply(this, args) as unknown;
  zzAdd("cacheStory", t);
  zzCount("normalizedBlocks", this.normalizedBlocks(args[0]));
  const blocks = args[1] as readonly unknown[];
  const list = (r ?? []) as readonly unknown[];
  zzCount(
    "replacedBlocks",
    list.length > 0 && Array.isArray(list[0])
      ? list.length
      : list.filter((block, index) => block !== blocks[index]).length,
  );
  return r as ReturnType<typeof zzStory>;
} as typeof zzStory;
'''
open(p,'w').write(s)
print('instrumented', p)
p=sys.argv[1]+'/packages/core/src/controller/canonicalSession.ts'
s=open(p).read()
a='    const normalized = normalizeCanonicalListRendering(\n      applied.value.document,\n      this.caches.lists,\n    );\n'
assert s.count(a)==1
s=s.replace(a,'    const zzNormalize = performance.now();\n'+a+'    zzAdd("normalizeInStage", zzNormalize);\n')
s+='''
function zzAdd(k: string, t: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + performance.now() - t;
}
'''
open(p,'w').write(s)
print('instrumented', p)
