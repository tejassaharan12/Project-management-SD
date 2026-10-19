import json,urllib.request,urllib.parse,os,sys
D=os.path.dirname(os.path.abspath(__file__)); T=open(D+'/.tok').read().strip(); UID='17841401712216851'
def get(fields,after=None):
    u=f'https://graph.facebook.com/v26.0/{UID}?'+urllib.parse.urlencode({'fields':fields,'access_token':T})
    try: return json.load(urllib.request.urlopen(u,timeout=60))
    except urllib.error.HTTPError as e: return {'error':json.load(e)}
H=['denverformen','blabliblu.official','bellavita.official','indewild','secret.alchemist','beardo.official','parkavenue_grooming','engagebyitc','ajmalperfumesin','house_of_em5','fraganote','plumgoodness']
MF='id,caption,timestamp,media_type,media_product_type,like_count,comments_count,permalink'
out={}
for h in H:
    posts=[];after=None;info=None
    while True:
        m=f'media.limit(50){{{MF}}}' if not after else f'media.limit(50).after({after}){{{MF}}}'
        r=get(f'business_discovery.username({h}){{username,name,followers_count,media_count,{m}}}')
        if 'error' in r: print(h,'ERROR',r['error'].get('error',{}).get('message')); break
        bd=r['business_discovery']; info=info or {k:bd.get(k) for k in ['username','name','followers_count','media_count']}
        md=bd.get('media',{}); data=md.get('data',[]); posts+=data
        after=md.get('paging',{}).get('cursors',{}).get('after')
        if not data or data[-1]['timestamp']<'2026-07-15' or not after: break
    posts=[p for p in posts if p['timestamp']>='2026-07-15']
    out[h]={'info':info,'posts':posts}; print(h,info and info['followers_count'],len(posts))
json.dump(out,open(D+'/ig_full.json','w'))
