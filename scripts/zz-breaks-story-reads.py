# Scratch only: applies each deliberate break of the in-place story reads and runs the block cache property.
import subprocess,pathlib,sys
pkg=pathlib.Path('packages/core')
F='src/prosemirror/conversion/toProseDoc.ts'
breaks=[
 ('none',None,None),
 ('advancing never updates the open comments','  if (reads.openCommentIds.end === reads.openCommentIds.start) return;','  return;'),
 ('the comparison ignores order','    if (ids[index] !== id) return false;','    if (!ids.includes(id)) return false;'),
 ('the end is always the start','        end: holdsInOrder(openStart, state.openCommentIds)','        end: true'),
 ('the comparison checks size only','  if (ids.length !== set.size) return false;\n  let index = 0;','  return ids.length === set.size;\n  let index = 0;'),
]
for name,old,new in breaks:
    p=pkg/F; s=p.read_text()
    if old is not None:
        assert s.count(old)==1,(name,s.count(old)); p.write_text(s.replace(old,new))
    try:
        r=subprocess.run(['bun','test','src/controller/projectionBlockCache.property.test.ts'],cwd=pkg,capture_output=True,text=True,timeout=1800)
        res=[l.strip() for l in (r.stdout+r.stderr).splitlines() if l.strip().endswith(' pass') or l.strip().endswith(' fail')]
        print(f'BREAK {name}: {" / ".join(res)}', flush=True)
    finally:
        p.write_text(s)
