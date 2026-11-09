# Scratch only: applies each deliberate break of the per-block story markers and runs the block cache property.
import subprocess,pathlib
pkg=pathlib.Path('packages/core')
F='src/prosemirror/conversion/toProseDoc.ts'
breaks=[
 ('none',None,None),
 ('ranged comments dropped','      for (const id of markers.commentStarts) ranged.add(id);\n',''),
 ('bookmarks counted out of order','      for (const [id, type] of markers.bookmarks) boundaries.count(id, type);','      for (const [id, type] of [...markers.bookmarks].reverse())\n        boundaries.count(id, type);'),
 ('a block with one kind kept as empty','          bookmarks.length === 0 && commentStarts.length === 0','          bookmarks.length === 0 || commentStarts.length === 0'),
 ('markers kept by the wrong block','      let markers = this.markers.get(block);','      let markers = this.markers.get(blocks[0] ?? block);'),
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
