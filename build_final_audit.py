#!/usr/bin/env python3
import glob,os,re,subprocess,html,json
from collections import Counter
from docx import Document
from docx.oxml.ns import qn
HERE=os.path.dirname(os.path.abspath(__file__))
BASE=os.path.expanduser("~/Desktop/Online Ed")
CUR=os.path.join(BASE,"Policies_Upload")
FIN=os.path.join(BASE,"Final Accessible Policy Uploads")
IMG=os.path.join(HERE,"img"); os.makedirs(IMG,exist_ok=True)
SOFFICE="/opt/homebrew/bin/soffice"; RES="100"
results={ (r['sub'],r['name']):r for r in json.load(open("/tmp/sf_results.json")) }
def esc(s): return html.escape(str(s),quote=True)
def code(n):
    m=re.match(r'^([A-Z]+)\s*(\d+)',n); return (m.group(1)+m.group(2)) if m else n[:8]
w=lambda t:Counter(re.findall(r"[a-z0-9]+",t.lower()))
IGN=set("cpc reviewed adopted approved endorsed board date next review page administrative procedure policy college of the canyons santa clarita spring fall revised".split())

# render final docs
tmp=os.path.join(HERE,"_pdf"); os.makedirs(tmp,exist_ok=True)
docs=[]
for sub in ("3000","4000","5000"):
    for f in sorted(glob.glob(os.path.join(FIN,sub,"*.docx"))):
        docs.append((sub,f))
files=[f for _,f in docs]
for k in range(0,len(files),20):
    subprocess.run([SOFFICE,"--headless","--convert-to","pdf","--outdir",tmp]+files[k:k+20],capture_output=True,timeout=600)
for sub,f in docs:
    c=code(os.path.basename(f))
    for old in glob.glob(os.path.join(IMG,c+"-*.png")): os.remove(old)
    pdf=os.path.join(tmp,os.path.splitext(os.path.basename(f))[0]+".pdf")
    if os.path.exists(pdf):
        subprocess.run(["pdftoppm","-png","-r",RES,pdf,os.path.join(IMG,c)],capture_output=True)
def pages(c):
    hits=glob.glob(os.path.join(IMG,c+"-*.png"))
    return [os.path.basename(x) for x in sorted(hits,key=lambda x:int(re.search(r'-(\d+)\.png$',x).group(1)))]

# per-doc metrics
def metrics(sub,f):
    d=Document(f)
    hs=[int(p.style.name.split()[1]) for p in d.paragraphs if p.style.name.startswith("Heading") and p.text.strip()]
    h1=hs.count(1); skip=any(hs[i]>hs[i-1]+1 for i in range(1,len(hs)))
    lang=d.styles.element.find(qn('w:docDefaults')) is not None
    alt=all((s._inline.find(qn('wp:docPr')) is not None and (s._inline.find(qn('wp:docPr')).get('descr') or '').strip()) for s in d.inline_shapes)
    acc = (h1==1 and not skip and lang and alt)
    # integrity: current docx tokens vs final rendered tokens
    cur=os.path.join(CUR,sub,os.path.basename(f))
    curtext=" ".join(p.text for p in Document(cur).paragraphs) if os.path.exists(cur) else ""
    c=code(os.path.basename(f))
    pdf=os.path.join(tmp,os.path.splitext(os.path.basename(f))[0]+".pdf")
    rt=subprocess.run(["pdftotext","-layout",pdf,"-"],capture_output=True,text=True).stdout if os.path.exists(pdf) else ""
    diff={k:v for k in set(w(curtext))|set(w(rt)) if (v:=w(curtext)[k]-w(rt)[k]) and k not in IGN and not k.isdigit()}
    r=results.get((sub,os.path.basename(f)),{})
    return dict(h=(hs.count(1),hs.count(2),hs.count(3)),acc=acc,skip=skip,lang=lang,alt=alt,
                integ=(not diff),rlists=r.get('real_lists',0),flags=r.get('flags',[]),
                irr=r.get('irregular_marker_paras',0))

secs={ "3000":[], "4000":[], "5000":[] }; toc={ "3000":[], "4000":[], "5000":[] }
n_acc=0; n_flag=0; n_int=0; tot_lists=0
for sub,f in docs:
    c=code(os.path.basename(f)); m=metrics(sub,f)
    n_acc+= 1 if m['acc'] else 0; n_int+= 1 if m['integ'] else 0
    tot_lists+=m['rlists']
    if m['flags'] or m['irr']: n_flag+=1
    toc[sub].append(f'<a href="#{c}" style="margin:2px 6px;color:#003366;text-decoration:none;font-size:12.5px;">{esc(c[:2]+" "+c[2:])}</a>')
    badges=(f'<span style="background:{"#1d6b3f" if m["acc"] else "#b3771a"};color:#fff;padding:2px 9px;border-radius:11px;font-size:11px;font-weight:700;">{"ACCESSIBLE" if m["acc"] else "REVIEW"}</span> '
            f'<span style="background:#eef;color:#334;padding:2px 8px;border-radius:11px;font-size:11px;">H1/H2/H3 = {m["h"][0]}/{m["h"][1]}/{m["h"][2]}</span> '
            f'<span style="background:#eef;color:#334;padding:2px 8px;border-radius:11px;font-size:11px;">{m["rlists"]} real lists</span> '
            f'<span style="background:{"#e8f3ec" if m["integ"] else "#fbeada"};color:#334;padding:2px 8px;border-radius:11px;font-size:11px;">integrity {"OK" if m["integ"] else "check"}</span>'
            + (f' <span style="background:#fff3d6;color:#8a5300;padding:2px 8px;border-radius:11px;font-size:11px;">&#9873; {len(m["flags"])+ (1 if m["irr"] else 0)} to review</span>' if (m['flags'] or m['irr']) else ''))
    flagnote=''
    if m['flags'] or m['irr']:
        items=m['flags'][:]
        if m['irr']: items.append(f"{m['irr']} paragraph(s) kept literal markers (paren/double-letter/single-item lists)")
        flagnote='<div style="background:#fff9ec;border-left:4px solid #d9a300;padding:8px 12px;margin:6px 0;font-size:12px;color:#6b5300;">Needs your eyes: '+'; '.join(esc(x) for x in items)+'</div>'
    imgs="".join(f'<img src="img/{p}" loading="lazy" style="width:100%;border:1px solid #ccd;display:block;margin-bottom:6px;">' for p in pages(c))
    secs[sub].append(f'<section id="{c}" style="margin:0 0 30px;border:1px solid #d7dee7;border-radius:9px;overflow:hidden;">'
        f'<div style="background:#003366;color:#fff;padding:10px 16px;position:sticky;top:0;z-index:5;"><b style="font-size:15px;">{esc(os.path.splitext(os.path.basename(f))[0])}</b>'
        f'<a href="#top" style="float:right;color:#cdd9e8;font-size:12px;text-decoration:none;">&#8679; top</a><div style="margin-top:6px;">{badges}</div></div>'
        f'<div style="padding:12px 16px;">{flagnote}{imgs}</div></section>')
import shutil; shutil.rmtree(tmp,ignore_errors=True)
N=len(docs)
banner=(f'<div style="background:#003366;color:#fff;border-radius:9px;padding:20px 24px;margin-bottom:14px;">'
 f'<div style="font-size:23px;font-weight:700;">Final Accessible Policy Uploads &mdash; Audit</div>'
 f'<div style="font-size:14px;color:#cdd9e8;margin-top:4px;">{N} policies &middot; real heading styles + real Word lists + accessibility tags &middot; College of the Canyons</div>'
 f'<div style="margin-top:12px;font-size:14px;">'
 f'<span style="background:#1d6b3f;padding:4px 12px;border-radius:12px;margin-right:6px;">{n_acc}/{N} pass accessibility</span>'
 f'<span style="background:#2c5a8a;padding:4px 12px;border-radius:12px;margin-right:6px;">{tot_lists} real lists</span>'
 f'<span style="background:#2c5a8a;padding:4px 12px;border-radius:12px;margin-right:6px;">{n_int}/{N} integrity clean</span>'
 f'<span style="background:#8a6a00;padding:4px 12px;border-radius:12px;">{n_flag} need review</span></div></div>')
tocblock="".join(f'<div style="background:#fff;border-radius:9px;padding:12px 18px;margin-bottom:10px;"><b style="color:#003366;">{s} series ({len(toc[s])})</b><div style="margin-top:6px;line-height:1.9;">{"".join(toc[s])}</div></div>' for s in ("3000","4000","5000"))
body="".join(secs["3000"]+secs["4000"]+secs["5000"])
page=f'''<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Final Accessible Policy Uploads - Audit</title></head>
<body style="margin:0;background:#eef1f5;font-family:'Segoe UI',Arial,sans-serif;color:#1a2733;" id="top"><div style="max-width:1000px;margin:0 auto;padding:22px 16px;">{banner}{tocblock}{body}</div></body></html>'''
open(os.path.join(HERE,"index.html"),"w").write(page)
print(f"built audit: {N} docs, accessible {n_acc}/{N}, integrity {n_int}/{N}, flagged {n_flag}")
