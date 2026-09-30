"""Request-local Closing calculation. No upload is persisted; no DB access here."""
import math
import re
from decimal import Decimal

import pandas as pd
from openpyxl import load_workbook

DEPOTS = {'so_bc22': 'Cempaka', 'so_bc13': 'Serang', 'so_bc66': 'Cilegon'}
SITE_CODES = {'BC22': 'Cempaka', 'BC13': 'Serang', 'BC66': 'Cilegon'}
CATEGORIES = ('Device', 'Macbook', 'ACC')
COMMON_COLUMNS = {
    'salesman': ('Salesman Name', 'Salesman', 'Sales Person', 'Salesperson Name'),
    'brand_name': ('Brand Name',), 'brand': ('Brand',),
    'group': ('Item Group Desc', 'Item Group Description', 'Item Group', 'Product Group'),
    'article_code': ('Article Code', 'Material', 'Material Code', 'Material Number'),
    'article': ('Article Description', 'Material Description', 'Description'),
}
BILLING_COLUMNS = {
    **COMMON_COLUMNS,
    'so': ('SO Number',),
    'amount': ('Total Nett Amount No Tax', 'Total Net Amount No Tax'),
    'site': ('Site Code', 'Site Desc', 'Depot', 'Depo'),
    'date': ('Billing Date',),
}
SO_COLUMNS = {
    **COMMON_COLUMNS,
    'so': ('Sales Order Number', 'SO Number'), 'do': ('DO Number', 'Delivery No'),
    'sales': ('TTL Sales Price',), 'discount': ('TTL Discount',),
    'qty': ('Quantity', 'Qty'), 'dealer': ('Soldto Name', 'Sold to Party Name', 'Sold to Name', 'Dealer'),
    'date': ('Order Date',),
    # Parsed for compatibility only; SO depot always comes from the upload slot.
    'site': ('Depot', 'Depo', 'Site Code', 'Site Desc', 'Plant', 'Plant Description'),
}


def clean(value):
    return '' if value is None or pd.isna(value) else str(value).strip()


def key(value):
    return re.sub(r'\s+', ' ', clean(value)).casefold()


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
                record = {name: row[index] if index is not None and index < len(row) else None
                          for name, index in mapping.items()}
                record['_has_brand'] = mapping['brand'] is not None or mapping['brand_name'] is not None
                record['_has_site'] = mapping['site'] is not None
                yield record
    finally:
        workbook.close()


def normalize_so(value):
    value = clean(value)
    if key(value) in ('', 'nan', 'none', 'null', '-', 'nat'):
        return ''
    if re.fullmatch(r'\d+(?:\.\d+)?[eE][+-]?\d+', value):
        value = format(Decimal(value), 'f')
    if re.fullmatch(r'\d+\.0+', value):
        value = value.split('.')[0]
    return value.lstrip('0') or ('0' if value else '')


def number(value, label, row_number, blank_zero=False):
    raw = clean(value)
    if not raw:
        if blank_zero:
            return 0.0
        raise ValueError(f'{label} baris {row_number}: nilai wajib kosong.')
    if isinstance(value, (int, float)):
        result = float(value)
    else:
        raw = re.sub(r'(?i)rp|\s', '', raw)
        if ',' in raw and '.' in raw:
            raw = raw.replace('.', '').replace(',', '.') if raw.rfind(',') > raw.rfind('.') else raw.replace(',', '')
        elif ',' in raw:
            raw = raw.replace(',', '.') if len(raw.rsplit(',', 1)[-1]) in (1, 2) else raw.replace(',', '')
        elif re.fullmatch(r'-?\d{1,3}(?:\.\d{3})+', raw):
            raw = raw.replace('.', '')
        try:
            result = float(raw)
        except ValueError:
            result = float('nan')
    if not math.isfinite(result):
        raise ValueError(f'{label} baris {row_number}: angka tidak valid ({value}).')
    return result


def code(value):
    return normalize_so(value)


# Anchored families, not arbitrary mentions (e.g. "case for iPhone").
# MBA, MB NEO and APP WATCH spellings were observed in the supplied exports.
APPLE_FAMILY = re.compile(
    r'^(?:(?:apple\s+)?(?:iphone|ipad|macbook|mac\s+mini|mac\s+studio|mac\s+pro|imac|airpods|earpods)\b'
    r'|apple\s+(?:pencil|watch)\b|(?:apple\s+)?magic\s+(?:mouse|keyboard|trackpad)\b'
    r'|mba\s+\d{2}\b|mb\s+neo\s+\d{2}\b|app\s+watch\b)', re.I)
NON_APPLE = re.compile(r'\b(?:samsung|galaxy|redmi|tecno|xiaomi|poco|oppo|vivo|infinix|realme|huawei|honor)\b', re.I)


def apple_mask(frame):
    """Brand is authoritative; blank/unknown brand is NOT description fallback."""
    names = frame['brand_name'].map(key)
    brands = frame['brand'].map(key)
    known_name = names.ne('')
    by_brand = names.eq('apple') | (~known_name & brands.isin(('apple', 'app')))
    # Reject conflicting non-Apple codes as well as non-Apple brand names.
    by_brand &= brands.isin(('', 'apple', 'app'))
    descriptions = frame['article'].map(clean)
    material = frame['article_code'].map(clean)
    by_description = (descriptions.str.match(APPLE_FAMILY) | material.str.match(APPLE_FAMILY))
    by_description &= ~(descriptions.str.contains(NON_APPLE) | material.str.contains(NON_APPLE))
    # A third-party compatibility claim is not evidence of Apple manufacture.
    compatibility = r'\b(?:compatible|compatibility|replacement|for|untuk)\b'
    by_description &= ~(descriptions.str.contains(compatibility, case=False) |
                        material.str.contains(compatibility, case=False))
    return by_brand.where(frame['_has_brand'], by_description)


def so_group_from_description(description):
    """Category inference is called ONLY after Apple identity is established."""
    name = key(description)
    # Accessories before devices, for explicitly branded accessory descriptions.
    if re.search(r'\b(pencil|keyboard|mouse|trackpad)\b', name):
        return 'Computer Accessories'
    if re.search(r'\b(case|cover|charger|adapter|cable|magsafe)\b', name):
        return 'Mobile Accessories'
    if re.search(r'\b(airpods|earpods)\b', name):
        return 'Audio'
    if re.search(r'\bwatch\b', name):
        return 'Wearable'
    if re.search(r'\bipad\b', name):
        return 'Tablet'
    if re.search(r'\biphone\b', name):
        return 'Mobile Phones'
    if re.search(r'\b(macbook|mba|mbp|mb neo|mac mini|mac studio|mac pro|imac)\b', name):
        return 'Computer'
    return ''


def dates_from_values(values):
    """Accept Excel serials, native datetime and day-first SAP export strings."""
    result = pd.Series(pd.NaT, index=values.index, dtype='datetime64[ns]')
    numeric = values.map(lambda v: isinstance(v, (int, float)) and not pd.isna(v))
    if numeric.any():
        result.loc[numeric] = pd.to_datetime(values[numeric].astype(float), unit='D', origin='1899-12-30', errors='coerce')
    remaining = ~numeric
    if remaining.any():
        # Explicit ISO handling avoids day-first swapping YYYY-MM-DD dates.
        text = values[remaining].map(clean)
        iso = text.str.match(r'^\d{4}-\d{2}-\d{2}')
        result.loc[text.index[iso]] = pd.to_datetime(text[iso], format='ISO8601', errors='coerce')
        result.loc[text.index[~iso]] = pd.to_datetime(text[~iso], dayfirst=True, format='mixed', errors='coerce')
    return result


def prepare_closing(billing_stream, so_streams, month, targets, classify, canonical_depo,
                    *, locked_salesmen, allowed_depos, canonical_salesman=None):
    """All frames live inside this request; scope comes from dashboard constants."""
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', month):
        raise ValueError('Periode Closing tidak valid.')
    if not so_streams:
        raise ValueError('Upload minimal satu file SO/DO.')
    normalize_owner = canonical_salesman or clean
    owners = {key(normalize_owner(v)) for v in locked_salesmen}

    def depot_of(value):
        text = clean(value)
        return SITE_CODES.get(text.upper(), canonical_depo(text))

    allowed = {depot_of(v) for v in allowed_depos} & set(DEPOTS.values())

    def scope(frame, slot_depot=None):
        if frame.empty:
            frame['depot'] = pd.Series(dtype=str)
            return frame
        # SO/DO: upload slot is authoritative, regardless of Plant/Site/Depot.
        # Billing keeps its existing source-site mapping.
        frame['depot'] = slot_depot if slot_depot is not None else frame['site'].map(depot_of)
        frame = frame.loc[frame['depot'].isin(allowed)].copy()
        frame = frame.loc[frame['salesman'].map(lambda v: key(normalize_owner(v))).isin(owners)].copy()
        return frame.loc[apple_mask(frame)].copy()

    billing = pd.DataFrame(list(rows_from_excel(
        billing_stream, BILLING_COLUMNS, ('so', 'amount', 'group', 'site', 'date', 'salesman'),
        'Billing Detail')), columns=[*BILLING_COLUMNS, '_has_brand', '_has_site'])
    billing = scope(billing)
    billing['category'] = billing['group'].map(classify)
    billing = billing.loc[billing['category'].isin(CATEGORIES)].copy()
    billing['normalized_so'] = billing['so'].map(normalize_so)
    # Match all scoped Billing SOs in this upload, irrespective of billing date.
    billing_so = set(billing.loc[billing['normalized_so'].ne(''), 'normalized_so'])
    article_groups = {code(a): g for a, g in zip(billing['article_code'], billing['group']) if code(a)}
    dates = dates_from_values(billing['date'])
    actual_rows = billing.loc[dates.dt.strftime('%Y-%m').eq(month)].copy()
    actual_rows['amount'] = [number(v, 'Billing Total Nett Amount No Tax', i + 2)
                             for i, v in enumerate(actual_rows['amount'])]
    actual = actual_rows.groupby(['depot', 'category'])['amount'].sum().to_dict()

    frames = []
    for slot, stream in so_streams.items():
        if slot not in DEPOTS:
            raise ValueError(f'Slot SO/DO tidak dikenal: {slot}.')
        depot = DEPOTS[slot]
        frame = pd.DataFrame(list(rows_from_excel(
            stream, SO_COLUMNS, ('so', 'sales', 'discount', 'article', 'salesman'), f'SO/DO {depot}')),
            columns=[*SO_COLUMNS, '_has_brand', '_has_site'])
        frame = scope(frame, depot)
        if frame.empty:
            continue
        # Retain existing monthly order-date behavior when dates are supplied.
        if frame['date'].map(clean).ne('').any():
            dates = dates_from_values(frame['date'])
            frame = frame.loc[dates.dt.strftime('%Y-%m').eq(month)].copy()
        if frame.empty:
            continue
        frame['group'] = [clean(g) or article_groups.get(code(a), '') or so_group_from_description(d)
                          for g, a, d in zip(frame['group'], frame['article_code'], frame['article'])]
        frame['category'] = frame['group'].map(classify)
        unknown = ~frame['category'].isin(CATEGORIES)
        if unknown.any():
            examples = ', '.join(frame.loc[unknown, 'article'].map(clean).iloc[:3])
            raise ValueError(f'SO/DO {depot}: kategori Apple tidak dapat ditentukan ({examples}). Tambahkan Item Group Desc.')
        frame['normalized_so'] = frame['so'].map(normalize_so)
        frame = frame.loc[frame['normalized_so'].ne('')].copy()
        frame['is_billed'] = frame['normalized_so'].isin(billing_so)
        frames.append(frame)
    so = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=[*SO_COLUMNS, 'depot', 'category', 'normalized_so', 'is_billed'])
    pending = so.loc[so['is_billed'].eq(False)].copy()
    pending['sales'] = [number(v, 'SO/DO TTL Sales Price', i + 2) for i, v in enumerate(pending['sales'])]
    pending['discount'] = [number(v, 'SO/DO TTL Discount', i + 2, True) for i, v in enumerate(pending['discount'])]
    pending['nett'] = pending['sales'] + pending['discount']
    pending_totals = pending.groupby(['depot', 'category'])['nett'].sum().to_dict()

    target = {}
    for row in targets:
        depot = depot_of(row.depo)
        if depot not in allowed or key(normalize_owner(row.salesman)) not in owners:
            continue
        for cat, field in (('Device', 'device_target'), ('Macbook', 'macbook_target'), ('ACC', 'acc_target')):
            target[(depot, cat)] = target.get((depot, cat), 0.0) + float(getattr(row, field) or 0)
    summary = {}
    for depot in ('ALL', *DEPOTS.values()):
        summary[depot] = {}
        for cat in (*CATEGORIES, 'TOTAL'):
            depots = tuple(DEPOTS.values()) if depot == 'ALL' else (depot,)
            cats = CATEGORIES if cat == 'TOTAL' else (cat,)
            billed = float(sum(actual.get((d, c), 0) for d in depots for c in cats))
            goal = float(sum(target.get((d, c), 0) for d in depots for c in cats))
            value = float(sum(pending_totals.get((d, c), 0) for d in depots for c in cats))
            estimated = billed + value
            summary[depot][cat] = dict(actual=billed, pending=value, estimated=estimated, target=goal,
                                      pct=(estimated / goal * 100 if goal else 0), gap=estimated - goal)
    details = []
    for row in pending.itertuples(index=False):
        do = normalize_so(row.do)
        details.append(dict(depot=row.depot, category=row.category, normalized_so=row.normalized_so,
                            so=clean(row.so), do=do or '-', dealer=clean(row.dealer), article=clean(row.article_code),
                            description=clean(row.article), qty=clean(row.qty), sales=float(row.sales),
                            discount=float(row.discount), nett=float(row.nett),
                            note='' if do else 'No DO belum tersedia'))
    return summary, details, len(so), len(pending)
