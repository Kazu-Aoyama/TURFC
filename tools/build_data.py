"""入力シート(xlsx)の表記ゆれを補正し，修正版xlsxとサイト用 data.js を作る。
使い方: python3 tools/build_data.py 入力.xlsx 出力.xlsx data.js data.js(画像対応表の読み込み元)
"""
import openpyxl,sys,re,json,datetime,unicodedata as U
from openpyxl.styles import Font,PatternFill,Alignment
src,dst,outjs=sys.argv[1:4]
wb=openpyxl.load_workbook(src)
LOG=[]
def hdr(ws): return {c.value:i for i,c in enumerate(ws[1]) if c.value}
def setv(ws,r,col,new,why):
    H=hdr(ws);c=ws.cell(row=r,column=H[col]+1);old=c.value
    if old==new: return
    c.value=new;LOG.append((ws.title,r,col,old,new,why))
def nf(v): return U.normalize('NFKC',v) if isinstance(v,str) else v
def rows(ws):
    H=hdr(ws)
    for r in range(2,ws.max_row+1):
        vals={k:ws.cell(row=r,column=i+1).value for k,i in H.items()}
        if any(v not in (None,'') for v in vals.values()): yield r,vals
KAN={'1':'一','2':'二','3':'三','4':'四'}
def fix_comp(v):
    if not isinstance(v,str): return v
    x=nf(v).strip()
    m=re.match(r'^国公立(?:大会)?\s*(\d)回戦$',x) or re.match(r'^国公立(?:大会)?\s*([一二三四])回戦$',x)
    if m: x='国公立大会 '+KAN.get(m.group(1),m.group(1))+'回戦'
    return x
VEN={'へボン':'ヘボン'}
def fix_ven(v):
    if not isinstance(v,str): return v
    x=nf(v).strip()
    for a,b in VEN.items(): x=x.replace(a,b)
    return x
for sh in ['試合','百年史_試合']:
    ws=wb[sh]
    for r,v in rows(ws):
        for col,f,why in [('チーム',nf,'全角英数字を半角に統一'),('相手',nf,'全角英数字を半角に統一（対戦成績の集計を一本化）'),('会場',fix_ven,'表記統一（全角→半角，「へボン」→「ヘボン」）'),('大会',fix_comp,'大会名の表記統一')]:
            if col in v and isinstance(v[col],str): setv(ws,r,col,f(v[col]),why)
ws=wb['対戦校']
for r,v in rows(ws): setv(ws,r,'相手',nf(v['相手']),'全角英数字を半角に統一')
PAT=re.compile(r'^(.*?)\s*\(([0-9０-９MＭ]*)\)\s*(?:\[\s*(.*?)\s*\])?\s*$')
for sh,gcol in [('試合メンバー','学年（当時）'),('百年史_試合メンバー','学年（当時）'),('相手メンバー','学年')]:
    ws=wb[sh]
    for r,v in rows(ws):
        g=v.get(gcol)
        if isinstance(g,str) and nf(g)!=g: setv(ws,r,gcol,nf(g),'全角数字を半角に統一');g=nf(g)
        if sh!='試合メンバー': continue
        n=v['氏名']
        if not isinstance(n,str): continue
        m=PAT.match(n)
        if m and ('[' in n or re.fullmatch(r'\d',nf(m.group(2)))):
            name,gr,st=m.group(1).strip(),nf(m.group(2)),m.group(3) or ''
            setv(ws,r,'氏名',name,'氏名欄に混入していた学年・得点を分離')
            if not g and gr: setv(ws,r,gcol,gr,'氏名欄の「(学年)」から転記')
            for num,k in re.findall(r'(\d+)\s*(PG|DG|T|G)',st):
                if not v.get(k): setv(ws,r,k,int(num),'氏名欄の「[得点]」から転記')
        elif re.search(r'（.+）$',n):
            setv(ws,r,'氏名',re.sub(r'\s*（.+）$','',n),'氏名欄に混入していた出身校を削除')
# log sheet
if '修正ログ' in wb.sheetnames: lg=wb['修正ログ']
else:
    lg=wb.create_sheet('修正ログ');lg.append(['シート','行','列','修正前','修正後','理由'])
for c in lg[1]: c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='00965E')
for x in LOG: lg.append(list(x))
for col,w in zip('ABCDEF',[16,7,14,34,34,40]): lg.column_dimensions[col].width=w
lg.freeze_panes='A2'
wb.save(dst)
from collections import Counter
print('changes',len(LOG),Counter((x[0],x[5]) for x in LOG).most_common())
# pack
def norm(v):
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d') if v.time()==datetime.time(0) else v.strftime('%Y-%m-%dT%H:%M')
    if isinstance(v,datetime.date): return v.strftime('%Y-%m-%d')
    if isinstance(v,datetime.time): return v.strftime('%H:%M')
    if isinstance(v,float) and v.is_integer(): return int(v)
    if isinstance(v,str) and v.startswith('='): return None
    return v
sheets={}
for ws in wb.worksheets:
    if ws.title in ('使い方','修正ログ'): continue
    cols=[c.value for c in ws[1]]
    while cols and cols[-1] is None: cols.pop()
    rs=[]
    for r in ws.iter_rows(min_row=2,max_col=len(cols),values_only=True):
        r=[norm(x) for x in r]
        if any(x not in (None,'') for x in r):
            while r and r[-1] is None: r.pop()
            rs.append(r)
    sheets[ws.title]={'cols':cols,'rows':rs}
L=[l for l in open(sys.argv[4],encoding='utf-8').read().split('\n') if l.startswith(('window.TURFC_LOCAL=','const LOCAL='))][0]
LOCAL=L[L.index('{'):].rstrip().rstrip(';')
json.loads(LOCAL)
with open(outjs,'w',encoding='utf-8') as f:
    f.write('/* 東京大学運動会ラグビー部 サイトデータ（入力シートから自動生成。直接編集しない） */\n')
    f.write('window.TURFC_DATA='+json.dumps({'__packed':1,'built':datetime.date.today().isoformat(),'sheets':sheets},ensure_ascii=False,separators=(',',':'))+';\n')
    f.write('window.TURFC_LOCAL='+LOCAL+';\n')
print('wrote',outjs)
