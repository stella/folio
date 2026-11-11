# usage: python3 -I incl.py <cpuprofile> <keys> name1 name2 ...: inclusive ms/key per function name (each sample counted once per name)
import json,sys
d=json.load(open(sys.argv[1])); keys=int(sys.argv[2]); names=sys.argv[3:]
nodes={n['id']:n for n in d['nodes']}; parent={}
for n in d['nodes']:
    for c in n.get('children',[]): parent[c]=n['id']
tot={k:0 for k in names}
for s,dt in zip(d['samples'],d['timeDeltas']):
    seen=set(); x=s
    while x is not None:
        f=nodes[x]['callFrame']['functionName']
        if f in tot and f not in seen: tot[f]+=dt; seen.add(f)
        x=parent.get(x)
for k in names: print(f"{tot[k]/1000/keys:7.2f} {k}")
