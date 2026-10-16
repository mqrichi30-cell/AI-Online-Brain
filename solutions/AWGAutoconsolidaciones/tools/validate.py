import json,sys,os,re
src,out,variant,tcf=sys.argv[1:5]
CATCH=json.load(open(tcf,encoding='utf-8-sig'))['properties']['definition']['actions']['Scope_-_Catch']
def conts(acts,path=()):
    yield path,acts
    for k,v in acts.items():
        if 'actions' in v: yield from conts(v['actions'],path+(k,))
        if 'else' in v: yield from conts(v['else'].get('actions',{}),path+(k,'else'))
        for c,cv in v.get('cases',{}).items(): yield from conts(cv.get('actions',{}),path+(k,c))
        if 'default' in v: yield from conts(v['default'].get('actions',{}),path+(k,'default'))
ok=True
def err(*a):
    global ok; ok=False; print('ERROR',*a)
for f in sorted(os.listdir(os.path.join(out,'Workflows'))):
    d=json.load(open(os.path.join(out,'Workflows',f),encoding='utf-8-sig'))
    o=json.load(open(os.path.join(src,'Workflows',f),encoding='utf-8-sig'))
    df=d['properties']['definition']; acts=df['actions']
    names={}
    for p,c in conts(acts):
        for k in c:
            if k in names: err(f,'dup',k)
            names[k]=c
            if len(k)>80: err(f,'long',k)
        for k,v in c.items():
            for r in (v.get('runAfter') or {}):
                if r not in c: err(f,'runAfter missing',k,'->',r)
        # cycle
        seen={}
        def dfs(n,stack):
            if n in stack: err(f,'cycle',n); return
            for r in (c[n].get('runAfter') or {}):
                if r in c: dfs(r,stack|{n})
        for k in c: dfs(k,frozenset())
    low={k.lower() for k in names}
    s=json.dumps(df)
    for m in set(re.findall(r"(?:outputs|body|actions|result|items)\('([^']+)'\)",s)):
        if m.lower() not in low: err(f,'ref to unknown action',m)
    for m in set(re.findall(r"variables\('([^']+)'\)",s)):
        if not any(v.get('type')=='InitializeVariable' and v['inputs']['variables'][0]['name']==m for v in acts.values()): err(f,'var not initialized at top',m)
    top=list(acts)
    if top[-2:]!=['Scope_-_Try','Scope_-_Catch']: err(f,'top order',top)
    if any(acts[k]['type']!='InitializeVariable' for k in top[:-2]): err(f,'non-var at top')
    if acts['Scope_-_Catch']!=CATCH: err(f,'catch modified')
    on=set(); [on.update(c.keys()) for p,c in conts(o['properties']['definition']['actions'])]
    missing=on-set(names)
    if missing-{'AWG_-_CON_-_03_HTTP_Send_Created_Consolidation_Message'}: err(f,'missing originals',missing)
    # sends
    sends=[]
    for p,c in conts(acts):
        for k,v in c.items():
            h=v.get('inputs',{}).get('host',{}) if isinstance(v.get('inputs'),dict) else {}
            op=h.get('operationId'); api=h.get('apiId','')
            if op in('SendEmailV2','SendEmailV3','SendMailWithOptions','StartAndWaitForAnApproval') or (op=='HttpRequest' and re.search(r'/(reply|send)$',v['inputs']['parameters']['Uri'])) or (op=='HttpRequest' and v['inputs']['parameters']['Uri']=='v1.0/me/messages'):
                sends.append((p,k,op,api.split('/')[-1],c))
    print('==',variant,f[:40],'top vars:',top[:-2])
    for p,k,op,api,c in sends:
        # gated by approval? walk preds transitively in container and parents
        def preds(c,n,acc):
            for r in (c[n].get('runAfter') or {}):
                if r not in acc: acc.add(r); preds(c,r,acc)
            return acc
        gate=[x for x in preds(c,k,set()) if x.endswith('Guard_Aprobado')]
        # parent containers
        anc=[x for x in p if 'Requiere' in x]
        print('   %-24s %-11s %-72s gate=%s' % (op,api,k,gate or anc or '-'))
    # smtp check
    if variant=='smtp':
        for p,k,op,api,c in sends:
            if api=='shared_office365': err(f,'outlook send remains',k)
print('OK' if ok else 'FAILED')
