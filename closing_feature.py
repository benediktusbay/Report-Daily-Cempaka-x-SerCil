"""Request-local Closing calculation. No uploaded data is persisted."""
import math
import re
import pandas as pd
from openpyxl import load_workbook

DEPOTS = {'so_bc22': 'Cempaka', 'so_bc13': 'Serang', 'so_bc66': 'Cilegon'}
CATEGORIES = ('Device', 'Macbook', 'ACC')
BILLING_COLUMNS = {
    'so': ('SO Number',), 'amount': ('Total Nett Amount No Tax', 'Total Net Amount No Tax'),
    'group': ('Item Group Desc', 'Item Group Description', 'Item Group'),
    'site': ('Site Code', 'Site Desc', 'Depot', 'Depo'),
    'date': ('Billing Date',), 'article_code': ('Article Code',),
    'article': ('Article Description',),
}
SO_COLUMNS = {
    'so': ('Sales Order Number', 'SO Number'), 'do': ('DO Number', 'Delivery No'),
    'sales': ('TTL Sales Price',), 'discount': ('TTL Discount',),
    'article_code': ('Article Code', 'Material'),
    'article': ('Article Description', 'Material Description', 'Description'),
    'group': ('Item Group Desc', 'Item Group Description', 'Item Group'),
    'qty': ('Quantity', 'Qty'), 'dealer': ('Soldto Name', 'Sold to Name', 'Dealer'),
    'date': ('Order Date',),
}

def key(value):
    return re.sub(r'\s+', ' ', str(value or '').strip()).casefold()

def columns(header, aliases, required, label):
    present = {key(value): i for i, value in enumerate(header) if value is not None}
    mapped = {name: next((present[key(alias)] for alias in choices if key(alias) in present), None)
              for name, choices in aliases.items()}
    missing = [aliases[name][0] for name in required if mapped[name] is None]
    if missing:
        raise ValueError(f'{label}: kolom wajib tidak ditemukan: {", ".join(missing)}.')
    return mapped

def rows_from_excel(stream, aliases, required, label):
    stream.seek(0)
    workbook = load_workbook(stream, read_only=True, data_only=True)
    try:
        sheet = workbook['Export'] if label == 'Billing Detail' and 'Export' in workbook else workbook.active
        iterator = sheet.iter_rows(values_only=True)
        mapping = None
        for _ in range(30):
            row = next(iterator, None)
            if row is None:
                break
            try:
                mapping = columns(row, aliases, required, label)
                break
            except ValueError:
                continue
        if mapping is None:
            raise ValueError(f'{label}: header tidak ditemukan. Perlu kolom: {", ".join(aliases[n][0] for n in required)}.')
        for row in iterator:
            if any(value is not None and str(value).strip() for value in row):
                yield {name: row[index] if index is not None and index < len(row) else None
                       for name, index in mapping.items()}
    finally:
        workbook.close()

def normalize_so(value):
    if value is None or pd.isna(value):
        return ''
    value = str(value).strip()
    if re.fullmatch(r'\d+\.0+', value):
        value = value.split('.')[0]
    if re.fullmatch(r'\d+(?:\.\d+)?[eE][+-]?\d+', value):
        value = format(float(value), '.0f')
    return value.lstrip('0') or ('0' if value and set(value) == {'0'} else value)

def number(value, label, row_number, blank_zero=False):
    if value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == '':
        if blank_zero:
            return 0.0
        raise ValueError(f'{label} baris {row_number}: nilai wajib kosong.')
    if isinstance(value, (int, float)):
        result = float(value)
        if math.isfinite(result):
            return result
        raise ValueError(f'{label} baris {row_number}: angka tidak valid ({value}).')
    raw = str(value).strip().replace('Rp', '').replace(' ', '')
    if ',' in raw and '.' in raw:
        raw = raw.replace('.', '').replace(',', '.') if raw.rfind(',') > raw.rfind('.') else raw.replace(',', '')
    elif ',' in raw:
        raw = raw.replace(',', '.') if len(raw.rsplit(',', 1)[-1]) in (1, 2) else raw.replace(',', '')
    elif raw.count('.') > 1:
        raw = raw.replace('.', '')
    try:
        result = float(raw)
        if math.isfinite(result):
            return result
        raise ValueError()
    except ValueError:
        raise ValueError(f'{label} baris {row_number}: angka tidak valid ({value}).') from None

def code(value):
    return normalize_so(value)

def so_group_from_description(description):
    """SO export lacks Item Group; infer only recognizable Apple product families."""
    name = key(description)
    if re.search(r'\b(ipad|tablet|tab\w*|megapad)\b', name):
        return 'Tablet'
    if re.search(r'\b(iphone|galaxy [sa]\w*|redmi|tecno|poco|xiaomi)\b', name):
        return 'Mobile Phones'
    if re.search(r'\b(macbook|mba|mbp|mb neo)\b', name):
        return 'Computer'
    if re.search(r'\b(airpods|earpods|earphone|headphone|buds|speaker)\b', name):
        return 'Audio'
    if re.search(r'\b(watch\w*|smart band|band)\b', name):
        return 'Wearable'
    if re.search(r'\b(pencil|keyboard|mouse|magic trackpad)\b', name):
        return 'Computer Accessories'
    if re.search(r'\b(case|cover|charger|adapter|cable|magsafe)\b', name):
        return 'Mobile Accessories'
    return ''

def prepare_closing(billing_stream, so_streams, month, targets, classify, canonical_depo):
    billing = list(rows_from_excel(billing_stream, BILLING_COLUMNS,
                                   ('so', 'amount', 'group', 'site', 'date'), 'Billing Detail'))
    billing_so = {normalize_so(row['so']) for row in billing if normalize_so(row['so'])}
    article_groups = {code(row['article_code']): row['group'] for row in billing
                      if code(row['article_code']) and classify(row['group']) in CATEGORIES}
    actual = {(depot, cat): 0.0 for depot in DEPOTS.values() for cat in CATEGORIES}
    site_codes = {'BC22': 'Cempaka', 'BC13': 'Serang', 'BC66': 'Cilegon'}
    for i, row in enumerate(billing, 2):
        date = pd.to_datetime(row['date'], errors='coerce')
        if pd.isna(date) or date.strftime('%Y-%m') != month:
            continue
        site = str(row['site'] or '').strip().upper()
        depot = site_codes.get(site, canonical_depo(row['site']))
        cat = classify(row['group'])
        if depot in DEPOTS.values() and cat in CATEGORIES:
            actual[(depot, cat)] += number(row['amount'], 'Billing Total Nett Amount No Tax', i)
    frames = []
    for slot, stream in so_streams.items():
        depot = DEPOTS[slot]
        records = list(rows_from_excel(stream, SO_COLUMNS, ('so', 'sales', 'discount', 'article'), f'SO/DO {depot}'))
        frame = pd.DataFrame(records)
        if frame.empty:
            continue
        frame['depot'] = depot
        frame['normalized_so'] = frame['so'].map(normalize_so)
        frame = frame[frame['normalized_so'] != ''].copy()
        if 'date' in frame and frame['date'].notna().any():
            dates = pd.to_datetime(frame['date'], errors='coerce')
            frame = frame[dates.dt.strftime('%Y-%m') == month].copy()
        if frame.empty:
            continue
        frame['is_billed'] = frame['normalized_so'].isin(billing_so)
        frame['sales'] = [number(v, f'SO/DO {depot} TTL Sales Price', i) for i, v in enumerate(frame['sales'], 2)]
        frame['discount'] = [number(v, f'SO/DO {depot} TTL Discount', i, True) for i, v in enumerate(frame['discount'], 2)]
        frame['nett'] = frame['sales'] + frame['discount']
        frame['group'] = [g if g and str(g).strip() else article_groups.get(code(a), '') or so_group_from_description(d)
                          for g, a, d in zip(frame['group'], frame['article_code'], frame['article'])]
        frame['category'] = frame['group'].map(classify)
        unknown = frame[frame['category'] == 'Other']
        if not unknown.empty:
            examples = ', '.join(str(v) for v in unknown['article'].head(3))
            raise ValueError(f'SO/DO {depot}: kategori tidak dapat ditentukan untuk {len(unknown)} baris ({examples}). Tambahkan Item Group Desc pada export atau periksa deskripsi produk.')
        frames.append(frame)
    if not frames:
        raise ValueError('Tidak ada baris SO/DO untuk periode yang dipilih.')
    so = pd.concat(frames, ignore_index=True)
    pending = so[~so['is_billed']].copy()
    target = {(depot, cat): 0.0 for depot in DEPOTS.values() for cat in CATEGORIES}
    for row in targets:
        depot = canonical_depo(row.depo)
        if depot in DEPOTS.values():
            target[(depot, 'Device')] += float(row.device_target or 0)
            target[(depot, 'Macbook')] += float(row.macbook_target or 0)
            target[(depot, 'ACC')] += float(row.acc_target or 0)
    summary = {}
    details = []
    for depot in ('ALL', *DEPOTS.values()):
        summary[depot] = {}
        for cat in (*CATEGORIES, 'TOTAL'):
            depots = tuple(DEPOTS.values()) if depot == 'ALL' else (depot,)
            cats = CATEGORIES if cat == 'TOTAL' else (cat,)
            billed = sum(actual[(d, c)] for d in depots for c in cats)
            goal = sum(target[(d, c)] for d in depots for c in cats)
            selection = pending[pending['depot'].isin(depots) & pending['category'].isin(cats)]
            value = float(selection['nett'].sum())
            estimated = billed + value
            summary[depot][cat] = dict(actual=billed, pending=value, estimated=estimated,
                                       target=goal, pct=(estimated / goal * 100 if goal else 0), gap=estimated - goal)
    for row in pending.itertuples(index=False):
        do = str(row.do).strip() if row.do is not None and not pd.isna(row.do) else ''
        details.append(dict(depot=row.depot, category=row.category, normalized_so=row.normalized_so,
                            so=str(row.so).strip(), do=do if do and do.lower() != 'nan' else '-',
                            dealer=str(row.dealer or ''), article=str(row.article_code or ''),
                            description=str(row.article or ''), qty=row.qty if row.qty is not None else '',
                            sales=row.sales, discount=row.discount, nett=row.nett,
                            note='' if do and do.lower() != 'nan' else 'No DO belum tersedia'))
    return summary, details, len(so), len(pending)
