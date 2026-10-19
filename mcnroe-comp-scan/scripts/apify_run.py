import json,sys,urllib.request,os,time
D=os.path.dirname(os.path.abspath(__file__)); T=open(D+'/.tok').read().strip()
H={'Authorization':'Bearer '+T,'Content-Type':'application/json'}
def req(u,data=None):
    return json.load(urllib.request.urlopen(urllib.request.Request(u,data=data,headers=H),timeout=120))
actor,name=sys.argv[1],sys.argv[2]; inp=sys.stdin.read()
maxc=sys.argv[3] if len(sys.argv)>3 else None
u=f'https://api.apify.com/v2/acts/{actor}/runs'+(f'?maxTotalChargeUsd={maxc}' if maxc else '')
r=req(u,inp.encode())['data']; rid=r['id']
while True:
    r=req(f'https://api.apify.com/v2/actor-runs/{rid}?waitForFinish=60')['data']
    if r['status'] not in ('READY','RUNNING'): break
items=req(f'https://api.apify.com/v2/datasets/{r["defaultDatasetId"]}/items?clean=true&format=json')
json.dump(items,open(f'{D}/{name}.json','w'))
u2=req('https://api.apify.com/v2/users/me/limits')['data']['current']['monthlyUsageUsd']
print(name,r['status'],len(items),'items, run cost $%.3f'%(r.get('usageTotalUsd') or 0),'| month usage $%.2f'%u2)
