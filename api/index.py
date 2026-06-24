from http.server import BaseHTTPRequestHandler
import json, re, io, os, base64, cgi
from urllib.request import urlopen, Request
from urllib.error import HTTPError

try:
    import pandas as pd
    HAS_PANDAS = True
except:
    HAS_PANDAS = False

GITHUB_TOKEN = os.environ.get('GITHUB_TOKEN', '')
GITHUB_REPO  = os.environ.get('GITHUB_REPO', '')
DATA_PATH    = 'data/dashboard.json'
BRANCH       = 'main'

def github_get():
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_PATH}"
    req = Request(url, headers={
        'Authorization': f'token {GITHUB_TOKEN}',
        'Accept': 'application/vnd.github.v3+json'
    })
    try:
        res = urlopen(req)
        body = json.loads(res.read())
        content = base64.b64decode(body['content']).decode('utf-8')
        return json.loads(content), body['sha']
    except HTTPError as e:
        if e.code == 404:
            return None, None
        raise

def github_put(data, sha=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_PATH}"
    content = base64.b64encode(
        json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
    ).decode('utf-8')
    payload = {
        'message': f"대시보드 업데이트: {data.get('year','?')}년 {data.get('month','?')}월",
        'content': content,
        'branch': BRANCH,
    }
    if sha:
        payload['sha'] = sha
    req = Request(url,
        data=json.dumps(payload).encode('utf-8'),
        headers={
            'Authorization': f'token {GITHUB_TOKEN}',
            'Accept': 'application/vnd.github.v3+json',
            'Content-Type': 'application/json'
        },
        method='PUT'
    )
    urlopen(req)

def fv(v, d=1):
    try: return round(float(v) * 10**d) / 10**d if v is not None and v != '' else 0
    except: return 0

def pv(v):
    try: return round(float(v) * 10000) / 100 if v is not None and v != '' else 0
    except: return 0

def extract(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    d = {}

    sA = pd.read_excel(xl, sheet_name='그룹사 매출', header=None)
    d['date']   = str(sA.iloc[12, 13] if pd.notna(sA.iloc[12, 13]) else '').strip()
    d['writer'] = str(sA.iloc[11, 13] if pd.notna(sA.iloc[11, 13]) else '').strip()
    dm = re.search(r'(\d{4})\.(\d{2})', d['date'])
    d['year']  = dm.group(1) if dm else '2026'
    d['month'] = dm.group(2) if dm else '05'

    sales_rows = {'대웅제약':4,'대웅바이오':5,'한올바이오':6,'디엔컴퍼니':7,'디엔코스메틱스':8,'그룹사 계':9}
    d['sales'] = {}
    for name, ri in sales_rows.items():
        r = sA.iloc[ri]
        p3 = [fv(r[3]), fv(r[4]), fv(r[5])]
        d['sales'][name] = {
            'prev_year': fv(r[2]), 'prev3_avg': round(sum(p3)/3*10)/10,
            'target': fv(r[6]), 'actual': fv(r[7]),
            'achieve': pv(r[8]), 'yoy': pv(r[9]), 'qoq': pv(r[10])
        }

    plA = pd.read_excel(xl, sheet_name='4 손익', header=None)
    plH = plA.iloc[2]
    mCols = {}
    for j, v in enumerate(plH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})년\s*0?(\d{1,2})월', str(v))
        if m: mCols[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    d['months'] = sorted(mCols.keys())
    d['monthly'] = {}
    for co, ri in {'대웅제약':3,'한올바이오':13,'대웅바이오':40,'시지바이오':49}.items():
        r = plA.iloc[ri]
        d['monthly'][co] = {m: fv(r[c]) for m, c in mCols.items()}

    retA = pd.read_excel(xl, sheet_name='2 환입', header=None)
    retH = retA.iloc[10]
    retMB = {}
    for j, v in enumerate(retH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: retMB[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    d['ret'] = {}
    for co, ri in {'대웅제약':12,'한올바이오':13,'디엔코스메틱스':14,'디엔컴퍼니':15,'대웅바이오':16,'시지바이오':17}.items():
        r = retA.iloc[ri]
        d['ret'][co] = {'trend': {}}
        for m, bc in retMB.items():
            v = r[bc+2]
            d['ret'][co]['trend'][m] = round(float(v)*10000)/100 if pd.notna(v) and float(v)!=0 else 0
        sm = sorted(retMB.keys())
        if sm:
            lm, bc = sm[-1], retMB[sm[-1]]
            d['ret'][co]['latest_rate']  = d['ret'][co]['trend'].get(lm, 0)
            d['ret'][co]['latest_sales'] = int(r[bc]) if pd.notna(r[bc]) else 0
            d['ret'][co]['latest_amt']   = fv(r[bc+1])
            d['ret'][co]['vs_prev_avg']  = round(float(r[bc+3])*10000)/100 if pd.notna(r[bc+3]) else 0
            d['ret'][co]['vs_prev3']     = round(float(r[bc+4])*10000)/100 if pd.notna(r[bc+4]) else 0

    dcA = pd.read_excel(xl, sheet_name='직거래 세부', header=None)
    yr_row, mo_row = dcA.iloc[1], dcA.iloc[2]
    dc_month_cols = {}
    cur_year = None
    for j in range(len(yr_row)):
        yv = yr_row.iloc[j]
        try:
            yi = int(float(yv))
            if yi in (2025, 2026): cur_year = str(yi)[2:]
        except: pass
        mv = mo_row.iloc[j]
        if pd.notna(mv) and cur_year:
            try:
                mo = int(float(mv))
                if 1 <= mo <= 12:
                    key = f"{cur_year}.{mo:02d}"
                    if key not in dc_month_cols:
                        dc_month_cols[key] = j
            except: pass
    r7dc = dcA.iloc[7]
    d['dc'] = {}
    for co, off in {'대웅제약':0,'대웅바이오':1,'한올바이오':2}.items():
        d['dc'][co] = {}
        for m, bc in dc_month_cols.items():
            v = r7dc.iloc[bc+off]
            d['dc'][co][m] = round(float(v)*1000)/10 if pd.notna(v) else 0

    dcB = pd.read_excel(xl, sheet_name='직거래 ', header=None)
    hdr10, hdr11 = dcB.iloc[10], dcB.iloc[11]
    grp_month_col = {}
    for j, v in enumerate(hdr10):
        if pd.notna(v):
            m = re.search(r'26년\s*(\d{1,2})월', str(v))
            if m:
                for k in range(j, min(j+10, len(hdr11))):
                    if '직거래율' in str(hdr11.iloc[k] if pd.notna(hdr11.iloc[k]) else ''):
                        grp_month_col[f"26.{int(m.group(1)):02d}"] = k
                        break
    for j, v in enumerate(hdr11):
        if pd.notna(v) and '직거래율' in str(v):
            for k in range(j, -1, -1):
                parent = str(hdr10.iloc[k]) if pd.notna(hdr10.iloc[k]) else ''
                m2 = re.search(r'(\d{2})\.(\d{1,2})월', parent)
                if m2:
                    key = f"{m2.group(1)}.{int(m2.group(2)):02d}"
                    if key not in grp_month_col:
                        grp_month_col[key] = j
                    break
    d['dc']['그룹사 계'] = {}
    r24 = dcB.iloc[24]
    for m, c in grp_month_col.items():
        v = r24.iloc[c]
        if pd.notna(v):
            d['dc']['그룹사 계'][m] = round(float(v)*1000)/10

    arA = pd.read_excel(xl, sheet_name='1 채권및회전일', header=None)
    arH = arA.iloc[5]
    arMC = {}
    for j, v in enumerate(arH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: arMC[f"{m.group(1)}.{int(m.group(2)):02d}"] = j+1
    d['ar'] = {}
    for ch, ri in {'대웅제약_종합도매':7,'대웅제약_간납도매':8,'대웅제약_수출':9,
                   '한올바이오_도매':14,'한올바이오_병원':15,
                   '대웅바이오_종합도매':21,'대웅바이오_간납도매':22,'시지바이오_내수':28}.items():
        r = arA.iloc[ri]
        d['ar'][ch] = {}
        for m, c in arMC.items():
            v = r.iloc[c]
            d['ar'][ch][m] = fv(v) if pd.notna(v) and float(v) < 900 else None

    rotA = pd.read_excel(xl, sheet_name='3회전관리', header=None)
    rotH = rotA.iloc[9]
    rotMs = {}
    for j, v in enumerate(rotH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: rotMs[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    srm = sorted(rotMs.keys())
    prevM = srm[-2] if len(srm) >= 2 else None
    currM = srm[-1] if srm else None
    d['rot'] = {}
    d['rot_months'] = {'prev': prevM, 'curr': currM}
    if prevM and currM:
        pb, cb = rotMs[prevM], rotMs[currM]
        for name, ri in {'대웅제약(오프라인)':11,'대웅제약(씽크)':12,'대웅제약(해외수출)':13,
                         '대웅제약(CH건기식)':14,'한올바이오':15,'시지바이오':16,'디엔컴퍼니':17,'합 계':18}.items():
            r = rotA.iloc[ri]
            d['rot'][name] = {
                'prev_delay_cnt': int(r.iloc[pb+2]) if pd.notna(r.iloc[pb+2]) else 0,
                'prev_delay_amt': fv(r.iloc[pb+3]),
                'curr_total_cnt': int(r.iloc[cb])   if pd.notna(r.iloc[cb])   else 0,
                'curr_total_amt': fv(r.iloc[cb+1]),
                'curr_delay_cnt': int(r.iloc[cb+2]) if pd.notna(r.iloc[cb+2]) else 0,
                'curr_delay_amt': fv(r.iloc[cb+3]),
            }

    budA = pd.read_excel(xl, sheet_name='7 그룹사예산', header=None)
    hdr28, hdr29 = budA.iloc[28], budA.iloc[29]
    yr26_start = None
    for j, v in enumerate(hdr28):
        if pd.notna(v) and '26' in str(v):
            yr26_start = j; break
    bud_month_cols = {}
    if yr26_start is not None:
        for j in range(yr26_start, len(hdr29)):
            v = str(hdr29.iloc[j]) if pd.notna(hdr29.iloc[j]) else ''
            m = re.search(r'영업예산\((\d{1,2})월\)', v)
            if m: bud_month_cols[f"{int(m.group(1))}월"] = j
    d['budget'] = {}
    for co, ri in {'대웅제약':31,'대웅바이오':32,'한올바이오':33,'CG바이오':34,'DNC':35,'합계':36}.items():
        r = budA.iloc[ri]
        d['budget'][co] = {'months': {}}
        for m, bc in bud_month_cols.items():
            d['budget'][co]['months'][m] = {
                'sales_budget': fv(r.iloc[bc], 2), 'sales_used': fv(r.iloc[bc+1], 2),
                'sales_remain': fv(r.iloc[bc+2], 2), 'sales_rate': pv(r.iloc[bc+3]),
            }

    d['badDebt'] = {
        'amount': '~5.7억', 'target': '100% 회수 / 신규 발생 ZERO',
        'done': [
            {'title':'대웅바이오 - 한국바이오제약 (5.5억)',
             'items':['실익여부 검토 및 회생계획안 인가 동의 진행','50% 현금 변제 및 50% 주식출자로 전액 회수 종결 완료']},
            {'title':'대웅제약 - ㈜영우의약품물류 (0.1억)',
             'items':['수금 계획 미 이행에 따른 독촉장 발송 및 담보 청구 절차 진행','담보 실행 전 채무자 구두 변제 협의 및 회수 종결']}
        ],
        'todo': [
            {'title':'그룹사 60일 이상 장기 미회수 채권 정리',
             'items':['5월 DNC 장기미회수 채권 6억 및 시지바이오 6개처 대손정리 지원 완료',
                      '대웅제약 B2C 파트 및 기타 법인 잔여 미회수 채권 정리 방안 논의 (회계팀)',
                      '회수불가채권 채널 재분류 진행']}
        ],
        'writer': 'DW 영업관리팀 박두환 · 2026.06.12'
    }
    return d


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200); self._cors(); self.end_headers()

    def do_POST(self):
        if not HAS_PANDAS:
            self._json(500, {'error': 'pandas not installed'}); return
        try:
            # cgi 모듈로 multipart 파싱
            ctype = self.headers.get('Content-Type', '')
            length = int(self.headers.get('Content-Length', 0))

            # environ 구성
            environ = {
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': ctype,
                'CONTENT_LENGTH': str(length),
            }
            fp = io.BytesIO(self.rfile.read(length))

            form = cgi.FieldStorage(fp=fp, headers=self.headers, environ=environ)

            file_item = form['file'] if 'file' in form else None
            if file_item is None or not file_item.filename:
                self._json(400, {'error': '파일 없음'}); return

            file_bytes = file_item.file.read()

            data = extract(file_bytes)
            _, sha = github_get()
            github_put(data, sha)

            self._json(200, {'ok': True, 'year': data['year'], 'month': data['month']})
        except Exception as e:
            import traceback
            self._json(500, {'error': str(e), 'detail': traceback.format_exc()})

    def do_GET(self):
        try:
            data, _ = github_get()
            if data:
                self._json(200, data)
            else:
                self._json(404, {'error': '데이터 없음'})
        except Exception as e:
            self._json(500, {'error': str(e)})

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code); self._cors()
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self.end_headers(); self.wfile.write(body)

    def log_message(self, *args): pass
