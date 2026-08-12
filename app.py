import os
import re
import math
import hashlib
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-me-in-production')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get(
    'DATABASE_URL', 'sqlite:///' + os.path.join(BASE_DIR, 'sales.db')
).replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

db = SQLAlchemy(app)

# Billing mapping agreed for the dashboard.
DEVICE_GROUPS = {'mobile phones', 'tablet'}
MAC_GROUPS = {'computer'}
ACC_GROUPS = {'audio', 'computer accessories', 'mobile accessories', 'tablets accessories', 'wearable'}
QVO_THRESHOLD = 46_000_000
WEEK_PCTS = {1: 0.75, 2: 0.90, 3: 1.00, 4: 1.00}
WEEK_END_DAY = {1: 7, 2: 14, 3: 21, 4: 31}
WEEK_START_DAY = {1: 1, 2: 8, 3: 15, 4: 22}
SKU_TARGETS = {'2-3': 13, '4-6': 6, '7-10': 5, '>10': 2}

# Canonical display names and aliases from Billing Detail.
SALESMAN_ALIASES = {
    'zefanya septania simorangkir': 'Zefanya Septania Simorangkir',
    'rafhyski alhasan': 'Rafhyski Alhasan',
    'ikmah novtianingrum': 'Ikmah Novtianingrum',
    'michael serafin sidik': 'Michael Serafin Sidik',
    'michael serafin sidik deactive': 'Michael Serafin Sidik',
    'andy varandy': 'Andy Varandy',
    'andy varandy deactive': 'Andy Varandy',
    'edi purnomo': 'Edi Purnomo',
    'edi suyitno': 'Edi Suyitno',
    'muhamad fajri': 'Muhamad Fajri',
    'muhamad fajri deactive': 'Muhamad Fajri',
}

# Fallback only when a BP is not found in Monthly Target. Ikmah intentionally has no
# fallback because Serang/Cilegon must be distinguished by BP from the monthly target.
SALESMAN_DEPO_FALLBACK = {
    'Zefanya Septania Simorangkir': 'Cempaka',
    'Rafhyski Alhasan': 'Cempaka',
    'Michael Serafin Sidik': 'Roxy',
    'Andy Varandy': 'Roxy',
    'Edi Purnomo': 'Tangerang',
    'Edi Suyitno': 'Tangerang',
    'Muhamad Fajri': 'Tangerang',
}

COLOR_WORDS = {
    'black','white','blue','green','red','yellow','purple','pink','orange','gray','grey','silver','gold','starlight',
    'midnight','natural','titanium','desert','graphite','space','rose','coral','teal','ultramarine','indigo','lavender',
    'navy','beige','brown','cream','mint','cyan','magenta','violet','jetblack','cosmic','sky','light','dark'
}


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='viewer')


# Legacy manual target table is retained so upgrading does not break an existing DB.
class Target(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), nullable=False)
    salesman = db.Column(db.String(160), nullable=False)
    device_target = db.Column(db.Float, default=0)
    macbook_target = db.Column(db.Float, default=0)
    acc_target = db.Column(db.Float, default=0)
    bo_target = db.Column(db.Integer, default=25)
    __table_args__ = (db.UniqueConstraint('month','salesman', name='uq_target_month_salesman'),)


class MonthlyTarget(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), nullable=False, index=True)
    depo = db.Column(db.String(80), nullable=False, index=True)
    bp = db.Column(db.String(80), nullable=False, index=True)
    dealer = db.Column(db.String(255), nullable=False)
    salesman = db.Column(db.String(160), nullable=False, index=True)
    device_target = db.Column(db.Float, default=0)
    macbook_target = db.Column(db.Float, default=0)
    acc_target = db.Column(db.Float, default=0)
    bo_target = db.Column(db.Integer, default=1)
    qvo_target = db.Column(db.Integer, default=1)
    __table_args__ = (db.UniqueConstraint('month','bp', name='uq_monthly_target_month_bp'),)


class UploadLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    uploaded_by = db.Column(db.String(80))
    rows_read = db.Column(db.Integer, default=0)
    rows_added = db.Column(db.Integer, default=0)


class TargetUploadLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    uploaded_by = db.Column(db.String(80))
    rows_read = db.Column(db.Integer, default=0)
    dealers_loaded = db.Column(db.Integer, default=0)


class Billing(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    row_hash = db.Column(db.String(64), unique=True, nullable=False)
    billing_date = db.Column(db.Date, nullable=False, index=True)
    salesman = db.Column(db.String(160), nullable=False, index=True)
    sold_to_code = db.Column(db.String(80), nullable=False, index=True)
    sold_to_name = db.Column(db.String(255), nullable=False)
    item_group = db.Column(db.String(160), nullable=False)
    article = db.Column(db.String(500), nullable=False)
    quantity = db.Column(db.Float, default=0)
    nett_amount = db.Column(db.Float, default=0)
    category = db.Column(db.String(20), nullable=False)
    sku_key = db.Column(db.String(300))


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        user = db.session.get(User, session['user_id'])
        if not user or user.role != 'admin':
            flash('Menu ini hanya untuk Admin.', 'danger')
            return redirect(url_for('dashboard'))
        return fn(*args, **kwargs)
    return wrapper


def normalize_col(s):
    return re.sub(r'[^a-z0-9]+', ' ', str(s).strip().lower()).strip()


def normalize_text(s):
    return re.sub(r'\s+', ' ', str(s or '').strip())


def canonical_salesman(name):
    clean = normalize_text(name)
    key = clean.lower()
    if key in ('', 'nan', 'none'):
        return ''
    if key in SALESMAN_ALIASES:
        return SALESMAN_ALIASES[key]
    return clean.title()


def canonical_depo(value):
    raw = normalize_text(value)
    key = raw.lower()
    for depo in ('Cempaka', 'Serang', 'Cilegon', 'Roxy', 'Tangerang'):
        if depo.lower() in key:
            return depo
    return raw.title() if raw else 'Unmapped'


def normalize_bp(value):
    if value is None:
        return ''
    s = str(value).strip()
    if s.lower() in ('nan', 'none', ''):
        return ''
    # Excel often exposes BP as 10046783.0; BP must be displayed without .0.
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    # Scientific notation fallback for numeric BP cells.
    try:
        if re.fullmatch(r'\d+(?:\.\d+)?[eE][+-]?\d+', s):
            s = format(float(s), '.0f')
    except Exception:
        pass
    return s


BILLING_ALIASES = {
    'billing_date': ['billing date','billingdate','bill date'],
    'salesman': ['salesman name','salesman','sales person','salesperson name'],
    'sold_to_code': ['sold to party code','sold-to party code','sold to code'],
    'sold_to_name': ['sold to party name','sold-to party name','sold to name'],
    'item_group': ['item group desc','item group description','item group'],
    'article': ['article description','article desc','article'],
    'quantity': ['quantity','qty'],
    'nett_amount': ['total nett amount with tax','total net amount with tax','nett amount with tax','net amount with tax']
}

TARGET_ALIASES = {
    'depo': ['depo', 'depot'],
    'bp': ['bp', 'sold to party code', 'sold to code'],
    'dealer': ['dealer name', 'dealer', 'sold to party name'],
    'salesman': ['sc name', 'salesman name', 'salesman'],
    'device_target': ['device total', 'device target', 'target device'],
    'acc_target': ['acc total', 'accessories total', 'acc target', 'target acc'],
    'macbook_target': ['mac total', 'macbook total', 'mac target', 'target mac'],
    'bo_target': ['target bo', 'bo target'],
    'qvo_target': ['target qvo', 'qvo target'],
}


def map_columns(columns, aliases, label='file'):
    normalized = {normalize_col(c): c for c in columns}
    mapped = {}
    for key, choices in aliases.items():
        for alias in choices:
            if normalize_col(alias) in normalized:
                mapped[key] = normalized[normalize_col(alias)]
                break
    missing = [k for k in aliases if k not in mapped]
    if missing:
        raise ValueError(f'Kolom wajib tidak ditemukan di {label}: ' + ', '.join(missing))
    return mapped


def classify(item_group):
    g = normalize_text(item_group).lower()
    if g in DEVICE_GROUPS:
        return 'Device'
    if g in MAC_GROUPS:
        return 'Macbook'
    if g in ACC_GROUPS:
        return 'ACC'
    return 'Other'


def sku_from_article(article, item_group):
    """
    Device SKU = model/type + storage.
    Everything after the FIRST storage token (GB/TB) is ignored.

    Examples:
      iPhone 17 Pro 1TB Silver        -> IPHONE 17 PRO 1TB
      iPhone 17 Pro 1TB Cosmic Orange -> IPHONE 17 PRO 1TB
      iPhone 17 Pro 256GB Deep Blue   -> IPHONE 17 PRO 256GB
      iPad Air 11 (M4) Wifi 128GB Blue -> IPAD AIR 11 (M4) WIFI 128GB

    This deliberately does not depend on a colour dictionary, so new colour
    names cannot accidentally create duplicate SKUs.
    """
    if normalize_text(item_group).lower() not in DEVICE_GROUPS:
        return None

    s = normalize_text(article)
    if not s or s.lower() in ('nan', 'none'):
        return None

    # Normalize separators/spaces but keep meaningful model punctuation such as (M4).
    s = re.sub(r'[/,_]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()

    # The SKU identity ends at the first storage token.
    storage_match = re.search(r'\b\d+(?:\.\d+)?\s*(?:GB|TB)\b', s, flags=re.I)
    if storage_match:
        base = s[:storage_match.end()]
        base = re.sub(
            r'\b(\d+(?:\.\d+)?)\s*(GB|TB)\b',
            lambda m: f"{m.group(1)}{m.group(2).upper()}",
            base,
            count=1,
            flags=re.I
        )
        return re.sub(r'\s+', ' ', base).strip().upper()

    # Defensive fallback for an unusual Device article without storage.
    # Strip known colour words, but normal Device SKU rows should use the path above.
    tokens = re.split(r'\s+', re.sub(r'[-]+', ' ', s))
    cleaned = []
    for token in tokens:
        bare = token.lower().strip('()[]{}.,')
        if bare in COLOR_WORDS:
            continue
        cleaned.append(token)
    base = re.sub(r'\s+', ' ', ' '.join(cleaned)).strip()
    return base.upper() if base else None


def row_hash(vals):
    raw = '|'.join(str(v) for v in vals)
    return hashlib.sha256(raw.encode('utf-8', errors='ignore')).hexdigest()


def month_range(month):
    start = pd.to_datetime(month + '-01').date()
    end = (pd.Timestamp(start) + pd.offsets.MonthEnd(0)).date()
    return start, end


def to_num(value, default=0):
    v = pd.to_numeric(value, errors='coerce')
    return default if pd.isna(v) else float(v)


def rupiah(x):
    try:
        return 'Rp {:,.0f}'.format(float(x)).replace(',', '.')
    except Exception:
        return 'Rp 0'


def rupiah_short(x):
    try:
        n = float(x or 0)
    except Exception:
        n = 0
    if abs(n) >= 1_000_000_000:
        txt = f'{n/1_000_000_000:.2f}'.rstrip('0').rstrip('.').replace('.', ',')
        return f'Rp{txt} M'
    if abs(n) >= 1_000_000:
        txt = f'{n/1_000_000:.1f}'.rstrip('0').rstrip('.').replace('.', ',')
        return f'Rp{txt} Jt'
    if abs(n) >= 1_000:
        txt = f'{n/1_000:.1f}'.rstrip('0').rstrip('.').replace('.', ',')
        return f'Rp{txt} Rb'
    return f'Rp{n:,.0f}'.replace(',', '.')


app.jinja_env.filters['rupiah'] = rupiah
app.jinja_env.filters['rupiah_short'] = rupiah_short


@app.before_request
def ensure_db():
    db.create_all()
    if User.query.count() == 0:
        username = os.environ.get('ADMIN_USERNAME', 'admin')
        password = os.environ.get('ADMIN_PASSWORD', 'admin123')
        db.session.add(User(username=username, password_hash=generate_password_hash(password), role='admin'))
        db.session.commit()


@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password_hash, password):
            session['user_id'] = user.id
            session['username'] = user.username
            session['role'] = user.role
            return redirect(url_for('dashboard'))
        flash('Username atau password salah.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


def target_lookup_for_month(month):
    targets = MonthlyTarget.query.filter_by(month=month).all()
    return targets, {normalize_bp(t.bp): t for t in targets}


def resolve_billing_owner(row, target_by_bp):
    bp = normalize_bp(row.sold_to_code)
    target = target_by_bp.get(bp)
    if target:
        return target.salesman, target.depo, target.dealer, target
    salesman = canonical_salesman(row.salesman)
    return salesman, SALESMAN_DEPO_FALLBACK.get(salesman, 'Unmapped'), normalize_text(row.sold_to_name), None


def matches_scope(salesman, depo, salesman_filter, depo_filter):
    if salesman_filter and salesman != salesman_filter:
        return False
    if depo_filter and depo != depo_filter:
        return False
    return True


def weekly_targets(bo_target):
    return {w: int(math.ceil((bo_target or 0) * WEEK_PCTS[w])) for w in range(1, 5)}


def sku_bucket(n):
    if n == 1:
        return '1'
    if 2 <= n <= 3:
        return '2-3'
    if 4 <= n <= 6:
        return '4-6'
    if 7 <= n <= 10:
        return '7-10'
    if n > 10:
        return '>10'
    return '0'


@app.route('/')
@login_required
def dashboard():
    latest = db.session.query(db.func.max(Billing.billing_date)).scalar()
    latest_target_month = db.session.query(db.func.max(MonthlyTarget.month)).scalar()
    default_month = latest.strftime('%Y-%m') if latest else (latest_target_month or datetime.now().strftime('%Y-%m'))
    month = request.args.get('month', default_month)
    depo_filter = request.args.get('depo','').strip()
    salesman_filter = request.args.get('salesman','').strip()
    start, end = month_range(month)

    monthly_targets, target_by_bp = target_lookup_for_month(month)
    billing_rows = Billing.query.filter(Billing.billing_date >= start, Billing.billing_date <= end).all()

    # Filter dropdown options come from target master + billing aliases.
    depos = sorted({t.depo for t in monthly_targets if t.depo and t.depo != 'Unmapped'})
    salesman_options = set()
    for t in monthly_targets:
        if not depo_filter or t.depo == depo_filter:
            salesman_options.add(t.salesman)
    for r in billing_rows:
        owner, depo, _, _ = resolve_billing_owner(r, target_by_bp)
        if owner and (not depo_filter or depo == depo_filter):
            salesman_options.add(owner)
    salesmen = sorted(salesman_options)
    if salesman_filter and salesman_filter not in salesmen:
        salesman_filter = ''

    # Target rows in the active scope.
    scoped_targets = [
        t for t in monthly_targets
        if matches_scope(t.salesman, t.depo, salesman_filter, depo_filter)
    ]

    # Dealer master starts with all monthly target dealers so zero-achievement dealers remain visible.
    dealer = {}
    for t in scoped_targets:
        key = normalize_bp(t.bp)
        dealer[key] = {
            'salesman': t.salesman, 'depo': t.depo, 'bp': key, 'dealer': t.dealer,
            'Device': 0.0, 'Macbook': 0.0, 'ACC': 0.0, 'skus': set(), 'last_date': None,
            'device_target': float(t.device_target or 0), 'macbook_target': float(t.macbook_target or 0),
            'acc_target': float(t.acc_target or 0), 'bo_target': int(t.bo_target or 0), 'qvo_target': int(t.qvo_target or 0),
            'is_target': True
        }

    scoped_billing = []
    for r in billing_rows:
        owner, depo, dealer_name, target = resolve_billing_owner(r, target_by_bp)
        if not matches_scope(owner, depo, salesman_filter, depo_filter):
            continue
        scoped_billing.append((r, owner, depo))
        bp = normalize_bp(r.sold_to_code)
        d = dealer.setdefault(bp, {
            'salesman': owner, 'depo': depo, 'bp': bp, 'dealer': dealer_name,
            'Device': 0.0, 'Macbook': 0.0, 'ACC': 0.0, 'skus': set(), 'last_date': None,
            'device_target': float(target.device_target or 0) if target else 0,
            'macbook_target': float(target.macbook_target or 0) if target else 0,
            'acc_target': float(target.acc_target or 0) if target else 0,
            'bo_target': int(target.bo_target or 0) if target else 0,
            'qvo_target': int(target.qvo_target or 0) if target else 0,
            'is_target': bool(target)
        })
        if r.category in ('Device','Macbook','ACC'):
            d[r.category] += float(r.nett_amount or 0)
        # Always recalculate Device SKU from the original Article Description.
        # This also corrects historical billing rows whose stored sku_key was
        # created before the latest SKU normalization rule.
        if r.category == 'Device':
            current_sku = sku_from_article(r.article, r.item_group)
            if current_sku:
                d['skus'].add(current_sku)
        if d['last_date'] is None or r.billing_date > d['last_date']:
            d['last_date'] = r.billing_date

    # Salesman target totals from Monthly Target.
    salesman_target = {}
    for t in scoped_targets:
        s = salesman_target.setdefault(t.salesman, {'device':0.0,'macbook':0.0,'acc':0.0,'bo':0,'qvo':0,'dealers':0})
        s['device'] += float(t.device_target or 0)
        s['macbook'] += float(t.macbook_target or 0)
        s['acc'] += float(t.acc_target or 0)
        s['bo'] += int(t.bo_target or 0)
        s['qvo'] += int(t.qvo_target or 0)
        s['dealers'] += 1

    salesman_actual = {}
    sku_detail = []
    dealer_detail = []
    active_dealers = 0
    for bp, d in dealer.items():
        total = d['Device'] + d['Macbook'] + d['ACC']
        bo = (d['Device'] + d['Macbook']) > 0
        qvo = total >= QVO_THRESHOLD
        sku_count = len(d['skus'])
        bucket = sku_bucket(sku_count)
        if total != 0:
            active_dealers += 1
        s = salesman_actual.setdefault(d['salesman'], {
            'device':0.0,'macbook':0.0,'acc':0.0,'bo':0,'qvo':0,
            'sku_bins':{'1':0,'2-3':0,'4-6':0,'7-10':0,'>10':0}
        })
        s['device'] += d['Device']; s['macbook'] += d['Macbook']; s['acc'] += d['ACC']
        if bo: s['bo'] += 1
        if qvo: s['qvo'] += 1
        if bucket in s['sku_bins']:
            s['sku_bins'][bucket] += 1
        if sku_count > 0:
            sku_detail.append({
                'salesman': d['salesman'], 'depo': d['depo'], 'bp': bp, 'dealer': d['dealer'],
                'sku_count': sku_count, 'bucket': bucket, 'sku_list': sorted(d['skus'])
            })
        dealer_detail.append({
            'salesman': d['salesman'], 'depo': d['depo'], 'bp': bp, 'dealer': d['dealer'],
            'device': d['Device'], 'macbook': d['Macbook'], 'acc': d['ACC'], 'total': total,
            'device_target': d['device_target'], 'macbook_target': d['macbook_target'], 'acc_target': d['acc_target'],
            'bo': bo, 'qvo': qvo, 'sku': sku_count, 'is_target': d['is_target']
        })

    all_people = sorted(set(salesman_target) | set(salesman_actual))
    table = []
    for salesman in all_people:
        a = salesman_actual.get(salesman, {'device':0,'macbook':0,'acc':0,'bo':0,'qvo':0,'sku_bins':{'1':0,'2-3':0,'4-6':0,'7-10':0,'>10':0}})
        t = salesman_target.get(salesman, {'device':0,'macbook':0,'acc':0,'bo':25,'qvo':0,'dealers':0})
        total = a['device'] + a['macbook'] + a['acc']
        target_total = t['device'] + t['macbook'] + t['acc']
        bo_target = t['bo'] or 25
        current_week = min(4, max(1, ((max([r.billing_date.day for r,owner,depo in scoped_billing if owner == salesman], default=1)-1)//7)+1))
        current_speed_target = weekly_targets(bo_target)[current_week]
        device_pct = a['device']/t['device']*100 if t['device'] else 0
        mac_pct = a['macbook']/t['macbook']*100 if t['macbook'] else 0
        acc_pct = a['acc']/t['acc']*100 if t['acc'] else 0
        bo_pct = a['bo']/bo_target*100 if bo_target else 0
        sales_pct = total/target_total*100 if target_total else 0
        bins = a['sku_bins']
        sku_score = min(bins['2-3'],13)+min(bins['4-6'],6)+min(bins['7-10'],5)+min(bins['>10'],2)
        sku_pct = sku_score/26*100
        health = 'Green' if a['bo'] >= current_speed_target else ('Amber' if a['bo'] >= max(1, math.ceil(current_speed_target*0.9)) else 'Red')
        table.append({
            'salesman':salesman,'device':a['device'],'macbook':a['macbook'],'acc':a['acc'],'total':total,
            'device_target':t['device'],'macbook_target':t['macbook'],'acc_target':t['acc'],'target_total':target_total,
            'device_pct':device_pct,'macbook_pct':mac_pct,'acc_pct':acc_pct,'sales_pct':sales_pct,
            'bo':a['bo'],'bo_target':bo_target,'bo_pct':bo_pct,'qvo':a['qvo'],'qvo_target':t['qvo'],
            'sku_bins':bins,'sku_score':sku_score,'sku_pct':sku_pct,'health':health,
            'overall_score':(bo_pct*0.45)+(sku_pct*0.30)+(min(sales_pct,150)*0.25)
        })

    leaderboard = sorted(table, key=lambda x:(x['overall_score'],x['bo'],x['qvo'],x['total']), reverse=True)
    for idx, x in enumerate(leaderboard,1): x['rank']=idx
    rank_map = {x['salesman']:x['rank'] for x in leaderboard}
    for x in table: x['rank']=rank_map.get(x['salesman'])
    table.sort(key=lambda x:x['rank'] or 999)

    # Speed Distribution per salesman and per week, cumulative BO. Future weeks are left blank.
    latest_in_scope = max([r.billing_date for r,_,_ in scoped_billing], default=None)
    speed_rows = []
    for x in table:
        wk_targets = weekly_targets(x['bo_target'])
        weekly = {}
        person_rows = [(r,owner,depo) for r,owner,depo in scoped_billing if owner == x['salesman']]
        for w in range(1,5):
            is_active = bool(latest_in_scope and latest_in_scope.day >= WEEK_START_DAY[w])
            if not is_active:
                weekly[w] = {'target':wk_targets[w], 'actual':None, 'pct':None, 'status':'Future'}
                continue
            cutoff_day = min(WEEK_END_DAY[w], end.day, latest_in_scope.day)
            cutoff = start.replace(day=cutoff_day)
            bo_by_bp = {}
            for r,owner,depo in person_rows:
                if r.billing_date > cutoff:
                    continue
                bp = normalize_bp(r.sold_to_code)
                vals = bo_by_bp.setdefault(bp, {'device':0.0,'macbook':0.0})
                if r.category == 'Device': vals['device'] += float(r.nett_amount or 0)
                elif r.category == 'Macbook': vals['macbook'] += float(r.nett_amount or 0)
            actual = sum(1 for v in bo_by_bp.values() if v['device']+v['macbook']>0)
            pct = actual/wk_targets[w]*100 if wk_targets[w] else 0
            weekly[w] = {'target':wk_targets[w], 'actual':actual, 'pct':pct, 'status':'On Track' if actual>=wk_targets[w] else 'Need Push'}
        speed_rows.append({'salesman':x['salesman'],'weeks':weekly})

    sku_rows = [{'salesman':x['salesman'], **x['sku_bins']} for x in table]

    cards = {
        'sales': sum(x['total'] for x in table),
        'sales_target': sum(x['target_total'] for x in table),
        'device': sum(x['device'] for x in table), 'device_target':sum(x['device_target'] for x in table),
        'macbook': sum(x['macbook'] for x in table), 'macbook_target':sum(x['macbook_target'] for x in table),
        'acc': sum(x['acc'] for x in table), 'acc_target':sum(x['acc_target'] for x in table),
        'bo': sum(x['bo'] for x in table), 'bo_target':sum(x['bo_target'] for x in table),
        'qvo': sum(x['qvo'] for x in table), 'qvo_target':sum(x['qvo_target'] for x in table),
        'dealers': active_dealers, 'target_dealers':len(scoped_targets), 'salesmen':len(table)
    }
    cards['sales_pct'] = cards['sales']/cards['sales_target']*100 if cards['sales_target'] else 0
    cards['bo_rate'] = cards['bo']/cards['bo_target']*100 if cards['bo_target'] else 0
    cards['qvo_rate'] = cards['qvo']/cards['qvo_target']*100 if cards['qvo_target'] else 0

    uploads = UploadLog.query.order_by(UploadLog.uploaded_at.desc()).limit(6).all()
    target_uploads = TargetUploadLog.query.order_by(TargetUploadLog.uploaded_at.desc()).limit(6).all()
    return render_template(
        'dashboard.html', month=month, depo_filter=depo_filter, salesman_filter=salesman_filter,
        depos=depos, salesmen=salesmen, cards=cards, table=table, leaderboard=leaderboard,
        speed_rows=speed_rows, sku_rows=sku_rows, sku_detail=sku_detail, dealer_detail=dealer_detail,
        sku_targets=SKU_TARGETS, uploads=uploads, target_uploads=target_uploads,
        qvo_threshold=QVO_THRESHOLD, latest_in_scope=latest_in_scope
    )


@app.route('/upload', methods=['POST'])
@admin_required
def upload():
    f = request.files.get('file')
    return_month = request.form.get('month', datetime.now().strftime('%Y-%m'))
    if not f or not f.filename:
        flash('Pilih file Billing Detail terlebih dahulu.', 'danger')
        return redirect(url_for('dashboard', month=return_month))
    if not f.filename.lower().endswith(('.xlsx','.xls')):
        flash('Format Billing Detail harus Excel (.xlsx/.xls).', 'danger')
        return redirect(url_for('dashboard', month=return_month))
    try:
        df = pd.read_excel(f, sheet_name='Export')
        cmap = map_columns(df.columns, BILLING_ALIASES, 'Billing Detail / sheet Export')
        added = 0
        for _, rr in df.iterrows():
            dt = pd.to_datetime(rr[cmap['billing_date']], errors='coerce')
            if pd.isna(dt):
                continue
            salesman_raw = normalize_text(rr[cmap['salesman']])
            code_raw = normalize_text(rr[cmap['sold_to_code']])
            code = normalize_bp(code_raw)
            name = normalize_text(rr[cmap['sold_to_name']])
            group = normalize_text(rr[cmap['item_group']])
            article = normalize_text(rr[cmap['article']])
            qty = to_num(rr[cmap['quantity']])
            amount = to_num(rr[cmap['nett_amount']])
            if not salesman_raw or not code or salesman_raw.lower() == 'nan':
                continue
            # Preserve raw code in the hash so existing V1 data does not duplicate after upgrade.
            h = row_hash([dt.date().isoformat(), salesman_raw, code_raw, name, group, article, qty, amount])
            if Billing.query.filter_by(row_hash=h).first():
                continue
            db.session.add(Billing(
                row_hash=h, billing_date=dt.date(), salesman=salesman_raw, sold_to_code=code,
                sold_to_name=name, item_group=group, article=article, quantity=qty, nett_amount=amount,
                category=classify(group), sku_key=sku_from_article(article, group)
            ))
            added += 1
        db.session.add(UploadLog(
            filename=secure_filename(f.filename), uploaded_by=session.get('username'),
            rows_read=len(df), rows_added=added
        ))
        db.session.commit()
        flash(f'Billing berhasil: {added} baris baru dari {len(df)} baris. Duplikat dilewati.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Upload Billing gagal: {e}', 'danger')
    return redirect(url_for('dashboard', month=return_month))


@app.route('/upload-target', methods=['POST'])
@admin_required
def upload_target():
    f = request.files.get('target_file')
    target_month = request.form.get('target_month', datetime.now().strftime('%Y-%m'))
    if not f or not f.filename:
        flash('Pilih file Target Bulanan terlebih dahulu.', 'danger')
        return redirect(url_for('dashboard', month=target_month))
    if not f.filename.lower().endswith(('.xlsx','.xls')):
        flash('Format Target harus Excel (.xlsx/.xls).', 'danger')
        return redirect(url_for('dashboard', month=target_month))
    try:
        df = pd.read_excel(f, sheet_name=0)
        cmap = map_columns(df.columns, TARGET_ALIASES, 'Target Bulanan')
        collapsed = {}
        for _, rr in df.iterrows():
            bp = normalize_bp(rr[cmap['bp']])
            if not bp:
                continue
            salesman = canonical_salesman(rr[cmap['salesman']])
            depo = canonical_depo(rr[cmap['depo']])
            dealer_name = normalize_text(rr[cmap['dealer']])
            if not salesman or not dealer_name:
                continue
            rec = collapsed.setdefault(bp, {
                'depo':depo,'bp':bp,'dealer':dealer_name,'salesman':salesman,
                'device_target':0.0,'macbook_target':0.0,'acc_target':0.0,'bo_target':0,'qvo_target':0
            })
            rec['device_target'] += to_num(rr[cmap['device_target']])
            rec['macbook_target'] += to_num(rr[cmap['macbook_target']])
            rec['acc_target'] += to_num(rr[cmap['acc_target']])
            rec['bo_target'] += int(round(to_num(rr[cmap['bo_target']])))
            rec['qvo_target'] += int(round(to_num(rr[cmap['qvo_target']])))

        if not collapsed:
            raise ValueError('Tidak ada dealer/BP valid yang dapat dibaca.')

        # Monthly target is a snapshot: re-uploading a month cleanly replaces that month.
        MonthlyTarget.query.filter_by(month=target_month).delete(synchronize_session=False)
        for rec in collapsed.values():
            db.session.add(MonthlyTarget(month=target_month, **rec))
        db.session.add(TargetUploadLog(
            month=target_month, filename=secure_filename(f.filename), uploaded_by=session.get('username'),
            rows_read=len(df), dealers_loaded=len(collapsed)
        ))
        db.session.commit()
        flash(f'Target {target_month} berhasil: {len(collapsed)} dealer/BP dimuat.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Upload Target gagal: {e}', 'danger')
    return redirect(url_for('dashboard', month=target_month))


@app.route('/admin/targets')
@admin_required
def targets():
    month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    rows = MonthlyTarget.query.filter_by(month=month).order_by(MonthlyTarget.depo, MonthlyTarget.salesman, MonthlyTarget.bp).all()
    summary = {}
    for t in rows:
        key = (t.depo,t.salesman)
        x = summary.setdefault(key, {'depo':t.depo,'salesman':t.salesman,'dealers':0,'device':0,'macbook':0,'acc':0,'bo':0,'qvo':0})
        x['dealers'] += 1; x['device'] += t.device_target or 0; x['macbook'] += t.macbook_target or 0; x['acc'] += t.acc_target or 0; x['bo'] += t.bo_target or 0; x['qvo'] += t.qvo_target or 0
    return render_template('targets.html', month=month, summary=sorted(summary.values(), key=lambda x:(x['depo'],x['salesman'])), target_count=len(rows))


@app.route('/admin/users', methods=['GET','POST'])
@admin_required
def users():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        role = request.form.get('role','viewer')
        if not username or not password:
            flash('Username dan password wajib diisi.', 'danger')
        elif User.query.filter_by(username=username).first():
            flash('Username sudah ada.', 'danger')
        else:
            db.session.add(User(username=username,password_hash=generate_password_hash(password),role=role))
            db.session.commit()
            flash('User berhasil dibuat.', 'success')
    return render_template('users.html', users=User.query.order_by(User.username).all())


@app.route('/api/dealers')
@login_required
def api_dealers():
    # Retained for backward compatibility; V2 dashboard already receives dealer data server-side.
    return jsonify([])


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=os.environ.get('FLASK_DEBUG') == '1')
