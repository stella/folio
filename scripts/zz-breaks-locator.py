# Scratch only: applies each deliberate break of the kept paragraph locator and runs its property.
import subprocess,pathlib
pkg=pathlib.Path('packages/docx-core')
F=pkg/'src/ops/paragraphLocator.ts'
tests=['src/ops/__tests__/paragraphLocator.property.test.ts']
breaks=[
 ('none',None,None),
 ('removed blocks keep their ids','  for (const block of removed)\n    for (const [key, located] of locationsIn(block)) {','  for (const block of removed.slice(0, 0))\n    for (const [key, located] of locationsIn(block)) {'),
 ('a location left where its block stood first','    : placed(location, body.content.indexOf(block));','    : location;'),
 ('the chain read only at its newest node','  for (let at: IndexNode | undefined = node; at !== undefined; at = at.parent) {\n    const held','  for (let at: IndexNode | undefined = node; at !== undefined; at = undefined) {\n    const held'),
 ('ids compared as spelled','  const key = idKey(blockId);\n  const held','  const key = blockId;\n  const held'),
 ('the suffix compared at the same index','    blocks[blocks.length - 1 - suffix] ===\n      previous[previous.length - 1 - suffix]','    blocks[blocks.length - 1 - suffix] ===\n      previous[blocks.length - 1 - suffix]'),
 ('a repeated id answered by its first holder','  if (held.length > 1 || block === undefined || located?.length !== 1)','  if (block === undefined || located === undefined || located.length === 0)'),
 ('a settled chain drops its oldest changes','  for (const at of chain.toReversed())','  for (const at of chain.slice(0, 1))'),
]
for name,old,new in breaks:
    s=F.read_text()
    if old is not None:
        assert s.count(old)==1,(name,s.count(old)); F.write_text(s.replace(old,new))
    try:
        r=subprocess.run(['bun','test',*tests],cwd=pkg,capture_output=True,text=True,timeout=1800)
        res=[l.strip() for l in (r.stdout+r.stderr).splitlines() if l.strip().endswith(' pass') or l.strip().endswith(' fail')]
        print(f'BREAK {name}: {" / ".join(res)}', flush=True)
    finally:
        F.write_text(s)
