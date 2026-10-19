import json,urllib.request,urllib.parse,os,csv,re
D=os.path.dirname(os.path.abspath(__file__)); K=open(D+'/.key').read().strip()
def api(ep,**p):
    p['key']=K; return json.load(urllib.request.urlopen('https://www.googleapis.com/youtube/v3/'+ep+'?'+urllib.parse.urlencode(p),timeout=30))
CH={'Denver':['UC098jLw1avgPp3r_EYqsdmw'],'BlaBliBlu':['UCzCa_eov4TBoZ-Kk4aUp_eA'],'Bellavita':['UC6mZXPjbDrjQFzXbiBHoaOA'],
'Inde Wild':['UCftmiFzEwMRyHat5iff7ZSQ'],'Secret Alchemist':['UC9hxVOq3xRCG26Jsbwd-3iw'],'Beardo':['UC8tE3gNWC-B6BqLuKkqivxw'],
'Park Avenue':['UC-ejtnYD7K4HufqjccfziXA'],'Engage':['UC7xMOMtXLrri5nif8pfrkSw'],'Ajmal':['UCCI1f7LB5IAgNd3L7IvQoPw','UC2nO1cSFUI61aKR4LVQpUHw'],
'EM5':['UCFENpB0si4hP_ToUJvUVnog'],'Fraganote':['UCa0CFYmtH6DCvv0w4r_vSFg'],'Plum':['UC6tQtwAuADVKuxIRKfwMZfw']}
# also look for BodyLovin channel
r=api('search',part='snippet',q='Plum BodyLovin',type='channel',maxResults=3,regionCode='IN')
print('bodylovin cands',[(i['id']['channelId'],i['snippet']['title']) for i in r['items']])
START='2026-06-01'
def dur(s):
    m=re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?',s or ''); h,mi,se=[int(x or 0) for x in m.groups()] if m else (0,0,0); return h*3600+mi*60+se
rows=[];chans={}
for b,ids in CH.items():
    for cid in ids:
        c=api('channels',part='contentDetails,statistics,snippet',id=cid)['items'][0]
        chans[cid]=dict(brand=b,title=c['snippet']['title'],subs=c['statistics'].get('subscriberCount'),videos=c['statistics'].get('videoCount'))
        up=c['contentDetails']['relatedPlaylists']['uploads']; tok=None; vids=[]; stop=False
        while not stop:
            p=dict(part='contentDetails',playlistId=up,maxResults=50)
            if tok:p['pageToken']=tok
            pl=api('playlistItems',**p)
            for it in pl['items']:
                d=it['contentDetails'].get('videoPublishedAt','')
                if d>=START: vids.append(it['contentDetails']['videoId'])
                elif d: stop=True
            tok=pl.get('nextPageToken')
            if not tok: break
        for i in range(0,len(vids),50):
            for v in api('videos',part='snippet,statistics,contentDetails',id=','.join(vids[i:i+50]))['items']:
                s=v['statistics'];sn=v['snippet']
                rows.append(dict(brand=b,channel=c['snippet']['title'],date=sn['publishedAt'][:10],id=v['id'],title=sn['title'],secs=dur(v['contentDetails'].get('duration')),
                  views=int(s.get('viewCount',0)),likes=int(s.get('likeCount',0) or 0),comments=int(s.get('commentCount',0) or 0),desc=sn.get('description','')[:400].replace('\n',' '),tags='|'.join(sn.get('tags',[])[:15])))
        print(b,cid,len(vids))
rows.sort(key=lambda r:(r['brand'],r['date']))
with open(D+'/videos.csv','w',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(rows[0].keys()));w.writeheader();w.writerows(rows)
json.dump(chans,open(D+'/channels.json','w'),indent=1)
print('total',len(rows))
