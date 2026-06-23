from http.server import BaseHTTPRequestHandler
import json, re, io, os

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from vercel_kv import kv
    USE_KV = True
except:
    USE_KV = False

KV_KEY = 'dashboard_data'
LOCAL_PATH = '/tmp/dashboard_data.json'

# ─────────────────────────────────────────────
# 검증된 파싱 로직 (참고사이트 37개 수치 전수 대조 통과)
# ─────────────────────────────────────────────
def fv(v, d=1):
    try: return round(float(v) * 10**d) / 10**d if v is not None and v != '' else 0
    except: return 0

def pv(v):
    try: return round(float(v) * 10000) / 100 if v is not None and v != '' else 0
    except: return 0

def extract(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    d = {}

    # ── 마감정보 ──
    sA = pd.read_excel(xl, sheet_name='그룹사 매출', header=None)
    d['date']   = str(sA.iloc[12, 13] if pd.notna(sA.iloc[12, 13]) else '').strip()
    d['writer'] = str(sA.iloc[11, 13] if pd.notna(sA.iloc[11, 13]) else '').strip()
    dm = re.search(r'(\d{4})\.(\d{2})', d['date'])
    d['year']  = dm.group(1) if dm else '2026'
    d['month'] = dm.group(2) if dm else '05'

    # ── 매출 KPI (그룹사 매출 시트) ──
    sales_rows = {'대웅제약':4,'대웅바이오':5,'한올바이오':6,'디엔컴퍼니':7,'디엔코스메틱스':8,'그룹사 계':9}
    d['sales'] = {}
    for name, ri in sales_rows.items():
        r = sA.iloc[ri]
        p3 = [fv(r[3]), fv(r[4]), fv(r[5])]
        d['sales'][name] = {
            'prev_year': fv(r[2]),
            'prev3_avg': round(sum(p3)/3*10)/10,
            'target': fv(r[6]), 'actual': fv(r[7]),
            'achieve': pv(r[8]), 'yoy': pv(r[9]), 'qoq': pv(r[10])
        }

    # ── 손익 월별 매출 (헤더 텍스트 기반) ──
    plA = pd.read_excel(xl, sheet_name='4 손익', header=None)
    plH = plA.iloc[2]
    mCols = {}
    for j, v in enumerate(plH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})년\s*0?(\d{1,2})월', str(v))
        if m: mCols[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    d['months'] = sorted(mCols.keys())
    pl_rows = {'대웅제약':3,'한올바이오':13,'대웅바이오':40,'시지바이오':49}
    d['monthly'] = {}
    for co, ri in pl_rows.items():
        r = plA.iloc[ri]
        d['monthly'][co] = {m: fv(r[c]) for m, c in mCols.items()}

    # ── 환입 (헤더 텍스트 기반) ──
    retA = pd.read_excel(xl, sheet_name='2 환입', header=None)
    retH = retA.iloc[10]
    retMB = {}
    for j, v in enumerate(retH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: retMB[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    ret_rows = {'대웅제약':12,'한올바이오':13,'디엔코스메틱스':14,'디엔컴퍼니':15,'대웅바이오':16,'시지바이오':17}
    d['ret'] = {}
    for co, ri in ret_rows.items():
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

    # ── 직거래율 ──
    # 법인별 추이: 직거래 세부 시트 row6 (헤더 기반)
    dcA = pd.read_excel(xl, sheet_name='직거래 세부', header=None)
    yr_row  = dcA.iloc[1]
    mo_row  = dcA.iloc[2]
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
    # pandas에서 그룹사 직거래율 = row7, month_col 자체가 대웅 col
    r7dc = dcA.iloc[7]
    d['dc'] = {}
    for co, off in {'대웅제약':0,'대웅바이오':1,'한올바이오':2}.items():
        d['dc'][co] = {}
        for m, bc in dc_month_cols.items():
            v = r7dc.iloc[bc+off]
            d['dc'][co][m] = round(float(v)*1000)/10 if pd.notna(v) else 0

    # 그룹사 계: 직거래  시트 row24 col14=당월 직거래율 (헤더 기반)
    dcB = pd.read_excel(xl, sheet_name='직거래 ', header=None)
    # row10에서 "26년 N월" 헤더 찾기, row11 서브헤더 "직거래율" col 찾기
    hdr10 = dcB.iloc[10]
    hdr11 = dcB.iloc[11]
    grp_month_col = {}
    for j, v in enumerate(hdr10):
        if pd.notna(v):
            m = re.search(r'26년\s*(\d{1,2})월', str(v))
            if m:
                # 이 컬럼 기준으로 서브헤더에서 "직거래율" 찾기
                for k in range(j, min(j+10, len(hdr11))):
                    sub = str(hdr11.iloc[k]) if pd.notna(hdr11.iloc[k]) else ''
                    if '직거래율' in sub:
                        grp_month_col[f"26.{int(m.group(1)):02d}"] = k
                        break
    # 과거 월(25년) 도 추가 - row11 직접 스캔
    hdr10b = dcB.iloc[10]
    for j, v in enumerate(hdr11):
        if pd.notna(v) and '직거래율' in str(v):
            # 상위 헤더에서 월 추정
            for k in range(j, -1, -1):
                parent = str(hdr10b.iloc[k]) if pd.notna(hdr10b.iloc[k]) else ''
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

    # ── 회전일 (헤더 텍스트 기반) ──
    arA = pd.read_excel(xl, sheet_name='1 채권및회전일', header=None)
    arH = arA.iloc[5]
    arMC = {}
    for j, v in enumerate(arH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: arMC[f"{m.group(1)}.{int(m.group(2)):02d}"] = j+1
    ar_rows = {
        '대웅제약_종합도매':7,'대웅제약_간납도매':8,'대웅제약_수출':9,
        '한올바이오_도매':14,'한올바이오_병원':15,
        '대웅바이오_종합도매':21,'대웅바이오_간납도매':22,
        '시지바이오_내수':28
    }
    d['ar'] = {}
    for ch, ri in ar_rows.items():
        r = arA.iloc[ri]
        d['ar'][ch] = {}
        for m, c in arMC.items():
            v = r.iloc[c]
            # 999 같은 이상값 제외
            d['ar'][ch][m] = fv(v) if pd.notna(v) and float(v) < 900 else None

    # ── 회전지연 (헤더 텍스트 기반) ──
    rotA = pd.read_excel(xl, sheet_name='3회전관리', header=None)
    rotH = rotA.iloc[9]
    rotMs = {}
    for j, v in enumerate(rotH):
        if pd.isna(v): continue
        m = re.match(r'(\d{2})\.(\d{1,2})월', str(v))
        if m: rotMs[f"{m.group(1)}.{int(m.group(2)):02d}"] = j
    srm = sorted(rotMs.keys())
    prevM, currM = (srm[-2] if len(srm)>=2 else None), (srm[-1] if srm else None)
    rot_rows = {
        '대웅제약(오프라인)':11,'대웅제약(씽크)':12,'대웅제약(해외수출)':13,
        '대웅제약(CH건기식)':14,'한올바이오':15,'시지바이오':16,'디엔컴퍼니':17,'합 계':18
    }
    d['rot'] = {}
    d['rot_months'] = {'prev': prevM, 'curr': currM}
    if prevM and currM:
        pb, cb = rotMs[prevM], rotMs[currM]
        for name, ri in rot_rows.items():
            r = rotA.iloc[ri]
            d['rot'][name] = {
                'prev_delay_cnt': int(r.iloc[pb+2]) if pd.notna(r.iloc[pb+2]) else 0,
                'prev_delay_amt': fv(r.iloc[pb+3]),
                'curr_total_cnt': int(r.iloc[cb])   if pd.notna(r.iloc[cb])   else 0,
                'curr_total_amt': fv(r.iloc[cb+1]),
                'curr_delay_cnt': int(r.iloc[cb+2]) if pd.notna(r.iloc[cb+2]) else 0,
                'curr_delay_amt': fv(r.iloc[cb+3]),
            }

    # ── 예산 (헤더 텍스트 기반) ──
    budA = pd.read_excel(xl, sheet_name='7 그룹사예산', header=None)
    # row28=연도, row29=월별 예산항목 헤더
    # 26년 영업예산 컬럼: row28에서 "26년" 찾고 row29에서 "영업예산(N월)" 찾기
    bud_rows = {'대웅제약':31,'대웅바이오':32,'한올바이오':33,'CG바이오':34,'DNC':35,'합계':36}
    hdr28 = budA.iloc[28]
    hdr29 = budA.iloc[29]
    # 26년 시작 컬럼 찾기
    yr26_start = None
    for j, v in enumerate(hdr28):
        if pd.notna(v) and '26' in str(v):
            yr26_start = j
            break
    bud_month_cols = {}
    if yr26_start is not None:
        for j in range(yr26_start, len(hdr29)):
            v = str(hdr29.iloc[j]) if pd.notna(hdr29.iloc[j]) else ''
            m = re.search(r'영업예산\((\d{1,2})월\)', v)
            if m:
                bud_month_cols[f"{int(m.group(1))}월"] = j
    d['budget'] = {}
    for co, ri in bud_rows.items():
        r = budA.iloc[ri]
        d['budget'][co] = {'months': {}}
        for m, bc in bud_month_cols.items():
            d['budget'][co]['months'][m] = {
                'sales_budget': fv(r.iloc[bc], 2),
                'sales_used':   fv(r.iloc[bc+1], 2),
                'sales_remain': fv(r.iloc[bc+2], 2),
                'sales_rate':   pv(r.iloc[bc+3]),
            }

    # ── 부실채권 텍스트 ──
    d['badDebt'] = {
        'amount': '~5.7억',
        'target': '100% 회수 / 신규 발생 ZERO',
        'done': [
            {'title':'대웅바이오 - 한국바이오제약 (5.5억)',
             'items':['대웅바이오 위수탁 채권으로 회생 진행 중','실익여부 검토 및 회생계획안 인가 동의 진행','50% 현금 변제 및 50% 주식출자로 전액 회수 종결 완료']},
            {'title':'대웅제약 - ㈜영우의약품물류 (0.1억)',
             'items':['수금 계획 미 이행에 따른 독촉장 발송 및 담보 청구 절차 진행','담보 실행 전 채무자 구두 변제 협의 및 회수 종결']}
        ],
        'todo': [
            {'title':'그룹사 60일 이상 장기 미회수 채권 정리',
             'items':['5월 DNC 장기미회수 채권 6억 및 시지바이오 6개처 대손정리 지원 완료','대웅제약 B2C 파트 및 기타 법인 잔여 미회수 채권 정리 방안 논의 (회계팀)','회수불가채권 채널 재분류 진행']}
        ],
        'writer': 'DW 영업관리팀 박두환 · 2026.06.12'
    }

    return d


# ─────────────────────────────────────────────
# HTTP 핸들러
# ─────────────────────────────────────────────
class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200); self._cors(); self.end_headers()

    def do_POST(self):
        if pd is None:
            self._json(500, {'error': 'pandas not installed'}); return
        try:
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)
            ct     = self.headers.get('Content-Type', '')
            boundary = ct.split('boundary=')[-1].encode()
            file_bytes = None
            for part in body.split(b'--' + boundary):
                if b'filename=' in part:
                    hdr_end = part.find(b'\r\n\r\n')
                    if hdr_end != -1:
                        file_bytes = part[hdr_end+4:].rstrip(b'\r\n--')
                        break
            if not file_bytes:
                self._json(400, {'error': '파일 없음'}); return

            data = extract(file_bytes)

            if USE_KV:
                kv.set(KV_KEY, json.dumps(data, ensure_ascii=False))
            else:
                with open(LOCAL_PATH, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False)

            self._json(200, {'ok': True, 'year': data['year'], 'month': data['month']})
        except Exception as e:
            self._json(500, {'error': str(e)})

    def do_GET(self):
        try:
            if USE_KV:
                raw  = kv.get(KV_KEY)
                data = json.loads(raw) if raw else None
            else:
                with open(LOCAL_PATH, encoding='utf-8') as f:
                    data = json.load(f)
            self._json(200, data)
        except:
            self._json(404, {'error': '데이터 없음'})

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
