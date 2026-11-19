"""Download the SALSE workbook from its OneDrive share link and refresh the data inside index.html."""
import base64, io, json, math, os, re, sys, urllib.request
import pandas as pd

XLSX = '/tmp/salse.xlsx'
UA = {'User-Agent': 'Mozilla/5.0'}


def fetch(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read(), r.geturl()


def download():
    link = os.environ['XLSX_URL'].strip()
    tries = []
    b64 = base64.urlsafe_b64encode(link.encode()).decode().rstrip('=')
    tries.append('https://api.onedrive.com/v1.0/shares/u!' + b64 + '/root/content')
    tries.append(link + ('&' if '?' in link else '?') + 'download=1')
    try:
        _, final = fetch(link)
        tries.append(final.replace('/redir?', '/download?').replace('/view.aspx?', '/download.aspx?'))
        tries.append(final + ('&' if '?' in final else '?') + 'download=1')
    except Exception as e:
        print('resolve failed:', e)
    for u in tries:
        try:
            data, _ = fetch(u)
            if data[:2] == b'PK':
                open(XLSX, 'wb').write(data)
                print('downloaded', len(data), 'bytes')
                return
            print('not an xlsx from', u[:60])
        except Exception as e:
            print('failed', u[:60], e)
    sys.exit('could not download the workbook - check the XLSX_URL secret (OneDrive link with "Anyone with the link")')


def num(v):
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


def s(v):
    return '' if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip()


def build():
    x = pd.ExcelFile(XLSX)
    sp = pd.read_excel(x, 'salsepov', header=None)
    sp['dt'] = pd.to_datetime(sp[1], errors='coerce')
    daily = []
    for i, r in sp[sp['dt'].notna()].iterrows():
        z = lambda c: round(num(r[c]) or 0, 2)
        d = r['dt'].strftime('%Y-%m-%d')
        daily.append({'date': d, 'month': d[:7], 'sales21': z(2), 'sales18': z(3), 'branch18': z(6), 'bag_mahmoud': z(7),
                      'branch21': z(8), 'bag_islam': z(9), 'bag_abueissa': z(10), '_i': i})
    last = max(k for k, d in enumerate(daily) if d['sales21'] or d['sales18'])
    daily = daily[:last + 1]

    targets = {}
    for d in daily:
        if d['date'][8:] != '01':
            continue
        row, prev = sp.loc[d['_i']], sp.loc[d['_i'] - 1] if d['_i'] - 1 in sp.index else None
        prev_is_date = prev is not None and pd.notna(prev['dt'])
        def pick(c):
            v = num(row[c])
            if v is not None and v > 100:
                return v
            p = num(prev[c]) if prev is not None else None
            return p if (not prev_is_date and p is not None and p > 100) else None
        targets[d['month']] = (pick(11), pick(12))

    F = ['sales21', 'sales18', 'branch18', 'bag_mahmoud', 'branch21', 'bag_islam', 'bag_abueissa']
    monthly = []
    for m in dict.fromkeys(d['month'] for d in daily):
        ds = [d for d in daily if d['month'] == m]
        o = {'month': m, **{f: round(sum(d[f] for d in ds), 2) for f in F}}
        o['target21'], o['target18'] = targets.get(m, (None, None))
        monthly.append(o)

    branch_nov = []
    hdr = sp.index[sp[24].astype(str).str.strip() == 'الاهداف'] if 24 in sp.columns else []
    if len(hdr):
        h = hdr[0]
        for i in range(h + 1, h + 8):
            r = sp.loc[i]
            if s(r[24]): branch_nov.append({'karat': '21', 'branch': s(r[24]), 'target': num(r[25]), 'sales': round(num(r[26]) or 0, 2)})
            if s(r[30]): branch_nov.append({'karat': '18', 'branch': s(r[30]), 'target': num(r[31]), 'sales': round(num(r[32]) or 0, 2)})
        for b in branch_nov:
            b['pct'] = round(b['sales'] / b['target'], 4) if b['target'] else None
        for b in branch_nov:
            if b['branch'] == 'الاسكندرية' and b['target']:
                monthly[-1]['target' + b['karat']] = b['target']

    MAP = {'JUN': '2025-06', 'JUL': '2025-07', 'AUG': '2025-08', 'SEP': '2025-09', 'OCT': '2025-10', 'NOV': '2025-11', 'DEC': '2025-12', 'JAN': '2026-01', 'FEB': '2026-02'}
    sale = pd.read_excel(x, 'SALE', header=None)
    heads = [s(h).upper() for h in sale.iloc[0]]
    clients = []
    for _, r in sale.iloc[1:].iterrows():
        c = s(r[0])
        if not c:
            continue
        o = {'client': c, **{m: 0 for m in MAP.values()}}
        for j, hh in enumerate(heads):
            if hh in MAP:
                o[MAP[hh]] = round(num(r[j]) or 0, 2)
        o['total'] = round(sum(o[m] for m in MAP.values()), 2)
        clients.append(o)
    clients.sort(key=lambda o: -o['total'])

    sl = pd.read_excel(x, 'salse', header=None)
    balances = []
    for _, r in sl.iloc[1:].iterrows():
        c = s(r[0])
        if not c:
            continue
        balances.append({'n': str(len(balances) + 1), 'client': c, 'area': s(r[6]) or 'غير محدد', 'credit_limit': round(num(r[1]) or 0, 2),
                         'gold21': round(num(r[2]) or 0, 2), 'cash': round(num(r[3]) or 0, 2), 'status': s(r[7])})

    branch_aug = None
    try:
        ta = pd.read_excel(x, 'تحليل الفروع', header=None)
        out = []
        for k, title in (('18', 'تحليل الفروع - عيار 18'), ('21', 'تحليل الفروع - عيار 21')):
            hs = [i for i, v in ta[0].items() if s(v).startswith(title)]
            if not hs:
                continue
            for i in range(hs[0] + 2, hs[0] + 9):
                r = ta.loc[i]
                if s(r[0]):
                    out.append({'karat': k, 'branch': s(r[0]), 'opening': num(r[2]), 'received': num(r[1]), 'available': num(r[4]), 'sold': num(r[3]), 'sell_through': num(r[5])})
        branch_aug = out or None
    except Exception as e:
        print('branch analysis skipped:', e)

    for d in daily:
        d.pop('_i')
    html = open('index.html', encoding='utf-8').read()
    m = re.search(r'window\.__DATA = (\{.*?\});</script>', html, re.S)
    old = json.loads(m.group(1))
    data = {'daily_sales': daily, 'monthly_sales': monthly, 'client_sales': clients, 'client_balances': balances,
            'branch_aug': branch_aug or old.get('branch_aug', []), 'branch_nov': branch_nov or old.get('branch_nov', [])}
    new_html = html[:m.start(1)] + json.dumps(data, ensure_ascii=False) + html[m.end(1):]
    if '#sync{display:none}' not in new_html:
        new_html = new_html.replace('</head>', '<style>#sync{display:none}</style>\n</head>', 1)
    open('index.html', 'w', encoding='utf-8').write(new_html)
    print('last sales day:', daily[-1]['date'], '| months:', len(monthly), '| clients:', len(clients), '| balances:', len(balances))


if __name__ == '__main__':
    {'download': download, 'build': build}[sys.argv[1]]()
