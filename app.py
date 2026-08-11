import os, re, hashlib
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
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///' + os.path.join(BASE_DIR, 'sales.db')).replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

db = SQLAlchemy(app)

DEVICE_GROUPS = {'mobile phones', 'tablet'}
MAC_GROUPS = {'computer'}
ACC_GROUPS = {'mobile accessories', 'computer accessories', 'tablets accessories', 'wearable', 'audio'}
QVO_THRESHOLD = 46_000_000
TARGET_BO = 25
WEEK_TARGETS = {1: 19, 2: 23, 3: 25, 4: 25, 5: 25}

COLOR_WORDS = {
    'black','white','blue','green','red','yellow','purple','pink','orange','gray','grey','silver','gold','starlight',
    'midnight','natural','titanium','desert','graphite','space','rose','coral','teal','ultramarine','indigo','lavender',
    'navy','beige','brown','cream','mint','cyan','magenta','violet'
}

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='viewer')

class Target(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), nullable=False)
    salesman = db.Column(db.String(160), nullable=False)
    device_target = db.Column(db.Float, default=0)
    macbook_target = db.Column(db.Float, default=0)
    acc_target = db.Column(db.Float, default=0)
    bo_target = db.Column(db.Integer, default=TARGET_BO)
    __table_args__ = (db.UniqueConstraint('month','salesman', name='uq_target_month_salesman'),)

class UploadLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    uploaded_by = db.Column(db.String(80))
    rows_read = db.Column(db.Integer, default=0)
    rows_added = db.Column(db.Integer, default=0)

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

ALIASES = {
    'billing_date': ['billing date','billingdate','bill date'],
    'salesman': ['salesman name','salesman','sales person','salesperson name'],
    'sold_to_code': ['sold to party code','sold-to party code','sold to code'],
    'sold_to_name': ['sold to party name','sold-to party name','sold to name'],
    'item_group': ['item group desc','item group description','item group'],
    'article': ['article description','article desc','article'],
    'quantity': ['quantity','qty'],
    'nett_amount': ['total nett amount with tax','total net amount with tax','nett amount with tax','net amount with tax']
}

def map_columns(columns):
    normalized = {normalize_col(c): c for c in columns}
    mapped = {}
    for key, aliases in ALIASES.items():
        for alias in aliases:
            a = normalize_col(alias)
            if a in normalized:
                mapped[key] = normalized[a]
                break
    missing = [k for k in ALIASES if k not in mapped]
    if missing:
        raise ValueError('Kolom wajib tidak ditemukan: ' + ', '.join(missing))
    return mapped


def classify(item_group):
    g = str(item_group).strip().lower()
    if g in DEVICE_GROUPS:
        return 'Device'
    if g in MAC_GROUPS:
        return 'Macbook'
    if g in ACC_GROUPS:
        return 'ACC'
    return 'Other'


def sku_from_article(article, item_group):
    if str(item_group).strip().lower() not in DEVICE_GROUPS:
        return None
    s = str(article).strip()
    # Keep type/model + storage, ignore common color words and trailing colour phrases.
    tokens = re.split(r'\s+', re.sub(r'[/,_-]+', ' ', s))
    cleaned = [t for t in tokens if t.lower().strip('()[]') not in COLOR_WORDS]
    storage_match = re.search(r'\b(\d+(?:\.\d+)?\s*(?:GB|TB))\b', s, flags=re.I)
    storage = re.sub(r'\s+', '', storage_match.group(1).upper()) if storage_match else ''
    base = ' '.join(cleaned)
    if storage:
        pos = re.search(re.escape(storage.replace(' ', '')), base.replace(' ', ''), flags=re.I)
    # remove duplicate storage/color noise, preserve enough model detail for stable SKU identity
    base = re.sub(r'\b(?:BLACK|WHITE|BLUE|GREEN|RED|YELLOW|PURPLE|PINK|ORANGE|GRAY|GREY|SILVER|GOLD|STARLIGHT|MIDNIGHT|NATURAL|TITANIUM|DESERT|GRAPHITE|ROSE|TEAL|ULTRAMARINE)\b', '', base, flags=re.I)
    base = re.sub(r'\s+', ' ', base).strip()
    return (base + (' | ' + storage if storage and storage not in base.upper().replace(' ', '') else '')).strip(' |')


def row_hash(vals):
    raw = '|'.join(str(v) for v in vals)
    return hashlib.sha256(raw.encode('utf-8', errors='ignore')).hexdigest()


def month_range(month):
    start = pd.to_datetime(month + '-01').date()
    end = (pd.Timestamp(start) + pd.offsets.MonthEnd(0)).date()
    return start, end


def rupiah(x):
    try:
        return 'Rp {:,.0f}'.format(float(x)).replace(',', '.')
    except Exception:
        return 'Rp 0'

app.jinja_env.filters['rupiah'] = rupiah

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

@app.route('/')
@login_required
def dashboard():
    latest = db.session.query(db.func.max(Billing.billing_date)).scalar()
    default_month = latest.strftime('%Y-%m') if latest else datetime.now().strftime('%Y-%m')
    month = request.args.get('month', default_month)
    salesman_filter = request.args.get('salesman','').strip()
    start, end = month_range(month)
    q = Billing.query.filter(Billing.billing_date >= start, Billing.billing_date <= end)
    rows = q.all()
    if salesman_filter:
        rows = [r for r in rows if r.salesman == salesman_filter]
    all_salesmen = sorted({r.salesman for r in Billing.query.all()})

    # dealer-level monthly aggregation per salesman
    dealer = {}
    for r in rows:
        key = (r.salesman, r.sold_to_code, r.sold_to_name)
        d = dealer.setdefault(key, {'Device':0,'Macbook':0,'ACC':0,'skus':set(),'last_date':r.billing_date})
        if r.category in ('Device','Macbook','ACC'):
            d[r.category] += float(r.nett_amount or 0)
        if r.sku_key and r.category == 'Device':
            d['skus'].add(r.sku_key)
        if r.billing_date > d['last_date']:
            d['last_date'] = r.billing_date

    salesman_stats = {}
    for (salesman, code, name), d in dealer.items():
        s = salesman_stats.setdefault(salesman, {'Device':0,'Macbook':0,'ACC':0,'bo':0,'qvo':0,'dealers':[], 'sku_bins':{'2-3':0,'4-6':0,'7-10':0,'>10':0}})
        s['Device'] += d['Device']; s['Macbook'] += d['Macbook']; s['ACC'] += d['ACC']
        bo = (d['Device'] + d['Macbook']) > 0
        qvo = (d['Device'] + d['Macbook'] + d['ACC']) >= QVO_THRESHOLD
        sku_count = len(d['skus'])
        if bo: s['bo'] += 1
        if qvo: s['qvo'] += 1
        if 2 <= sku_count <= 3: s['sku_bins']['2-3'] += 1
        elif 4 <= sku_count <= 6: s['sku_bins']['4-6'] += 1
        elif 7 <= sku_count <= 10: s['sku_bins']['7-10'] += 1
        elif sku_count > 10: s['sku_bins']['>10'] += 1
        s['dealers'].append({'code':code,'name':name,**d,'bo':bo,'qvo':qvo,'sku_count':sku_count})

    targets = {t.salesman:t for t in Target.query.filter_by(month=month).all()}
    cards = {'sales':0,'bo':0,'qvo':0,'dealers':len(dealer)}
    table = []
    for salesman in sorted(salesman_stats):
        s = salesman_stats[salesman]
        t = targets.get(salesman)
        dev_t = t.device_target if t else 0
        mac_t = t.macbook_target if t else 0
        acc_t = t.acc_target if t else 0
        bo_t = t.bo_target if t else TARGET_BO
        total = s['Device']+s['Macbook']+s['ACC']
        cards['sales'] += total; cards['bo'] += s['bo']; cards['qvo'] += s['qvo']
        table.append({
            'salesman':salesman,'device':s['Device'],'macbook':s['Macbook'],'acc':s['ACC'],'total':total,
            'device_target':dev_t,'macbook_target':mac_t,'acc_target':acc_t,'bo_target':bo_t,
            'device_pct':(s['Device']/dev_t*100 if dev_t else 0),'macbook_pct':(s['Macbook']/mac_t*100 if mac_t else 0),'acc_pct':(s['ACC']/acc_t*100 if acc_t else 0),
            'bo':s['bo'],'bo_pct':s['bo']/bo_t*100 if bo_t else 0,'qvo':s['qvo'], 'sku_bins':s['sku_bins']
        })

    latest_in_month = max([r.billing_date for r in rows], default=start)
    week = min(5, ((latest_in_month.day - 1)//7)+1)
    speed_target = WEEK_TARGETS[week]
    speed = [{'salesman':x['salesman'],'bo':x['bo'],'target':speed_target,'pct':x['bo']/speed_target*100 if speed_target else 0,'status':'On Track' if x['bo']>=speed_target else 'Need Push'} for x in table]
    sku = [{'salesman':x['salesman'],**x['sku_bins']} for x in table]

    # Management view helpers
    for x in table:
        target_total = x['device_target'] + x['macbook_target'] + x['acc_target']
        x['sales_pct'] = (x['total'] / target_total * 100) if target_total else 0
        x['target_total'] = target_total
        bins = x['sku_bins']
        sku_score = min(bins['2-3'], 13) + min(bins['4-6'], 6) + min(bins['7-10'], 5) + min(bins['>10'], 2)
        x['sku_score'] = sku_score
        x['sku_pct'] = sku_score / 26 * 100
        x['overall_score'] = (x['bo_pct'] * 0.45) + (x['sku_pct'] * 0.30) + (min(x['sales_pct'], 150) * 0.25)
        x['health'] = 'Green' if x['bo'] >= speed_target else ('Amber' if x['bo'] >= max(1, speed_target-3) else 'Red')

    leaderboard = sorted(table, key=lambda x: (x['overall_score'], x['bo'], x['qvo'], x['total']), reverse=True)
    for rank, x in enumerate(leaderboard, start=1):
        x['rank'] = rank
    rank_map = {x['salesman']: x['rank'] for x in leaderboard}
    for x in table:
        x['rank'] = rank_map.get(x['salesman'])

    cards['device'] = sum(x['device'] for x in table)
    cards['macbook'] = sum(x['macbook'] for x in table)
    cards['acc'] = sum(x['acc'] for x in table)
    cards['on_track'] = sum(1 for x in speed if x['status'] == 'On Track')
    cards['salesmen'] = len(table)
    cards['qvo_rate'] = (cards['qvo'] / cards['dealers'] * 100) if cards['dealers'] else 0
    cards['bo_rate'] = (cards['bo'] / (len(table) * TARGET_BO) * 100) if table else 0

    uploads = UploadLog.query.order_by(UploadLog.uploaded_at.desc()).limit(8).all()
    return render_template('dashboard.html', month=month, salesman_filter=salesman_filter, salesmen=all_salesmen,
                           cards=cards, table=table, leaderboard=leaderboard, speed=speed, sku=sku, week=week, speed_target=speed_target, uploads=uploads)

@app.route('/upload', methods=['POST'])
@admin_required
def upload():
    f = request.files.get('file')
    if not f or not f.filename:
        flash('Pilih file Excel terlebih dahulu.', 'danger'); return redirect(url_for('dashboard'))
    if not f.filename.lower().endswith(('.xlsx','.xls')):
        flash('Format harus Excel (.xlsx/.xls).', 'danger'); return redirect(url_for('dashboard'))
    try:
        df = pd.read_excel(f, sheet_name='Export')
        cmap = map_columns(df.columns)
        added = 0
        for _, rr in df.iterrows():
            dt = pd.to_datetime(rr[cmap['billing_date']], errors='coerce')
            if pd.isna(dt):
                continue
            salesman = str(rr[cmap['salesman']]).strip()
            code = str(rr[cmap['sold_to_code']]).strip()
            name = str(rr[cmap['sold_to_name']]).strip()
            group = str(rr[cmap['item_group']]).strip()
            article = str(rr[cmap['article']]).strip()
            qty = pd.to_numeric(rr[cmap['quantity']], errors='coerce')
            amount = pd.to_numeric(rr[cmap['nett_amount']], errors='coerce')
            qty = 0 if pd.isna(qty) else float(qty)
            amount = 0 if pd.isna(amount) else float(amount)
            if salesman.lower() == 'nan' or code.lower() == 'nan':
                continue
            h = row_hash([dt.date().isoformat(),salesman,code,name,group,article,qty,amount])
            if Billing.query.filter_by(row_hash=h).first():
                continue
            cat = classify(group)
            db.session.add(Billing(row_hash=h,billing_date=dt.date(),salesman=salesman,sold_to_code=code,sold_to_name=name,
                                   item_group=group,article=article,quantity=qty,nett_amount=amount,category=cat,sku_key=sku_from_article(article,group)))
            added += 1
        db.session.add(UploadLog(filename=secure_filename(f.filename), uploaded_by=session.get('username'), rows_read=len(df), rows_added=added))
        db.session.commit()
        flash(f'Upload berhasil: {added} baris baru dari {len(df)} baris. Data duplikat otomatis dilewati.', 'success')
    except Exception as e:
        db.session.rollback(); flash(f'Upload gagal: {e}', 'danger')
    return redirect(url_for('dashboard'))

@app.route('/admin/targets', methods=['GET','POST'])
@admin_required
def targets():
    month = request.values.get('month', datetime.now().strftime('%Y-%m'))
    if request.method == 'POST':
        salesman = request.form['salesman'].strip()
        t = Target.query.filter_by(month=month, salesman=salesman).first() or Target(month=month, salesman=salesman)
        t.device_target = float(request.form.get('device_target') or 0)
        t.macbook_target = float(request.form.get('macbook_target') or 0)
        t.acc_target = float(request.form.get('acc_target') or 0)
        t.bo_target = int(request.form.get('bo_target') or TARGET_BO)
        db.session.add(t); db.session.commit(); flash('Target tersimpan.', 'success')
    salesmen = sorted({r.salesman for r in Billing.query.all()})
    target_rows = {t.salesman:t for t in Target.query.filter_by(month=month).all()}
    return render_template('targets.html', month=month, salesmen=salesmen, targets=target_rows)

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
            db.session.add(User(username=username,password_hash=generate_password_hash(password),role=role)); db.session.commit(); flash('User berhasil dibuat.', 'success')
    return render_template('users.html', users=User.query.order_by(User.username).all())

@app.route('/api/dealers')
@login_required
def api_dealers():
    month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    salesman = request.args.get('salesman','')
    start,end = month_range(month)
    rows = Billing.query.filter(Billing.billing_date>=start,Billing.billing_date<=end).all()
    if salesman: rows=[r for r in rows if r.salesman==salesman]
    agg={}
    for r in rows:
        k=(r.salesman,r.sold_to_code,r.sold_to_name)
        d=agg.setdefault(k,{'device':0,'macbook':0,'acc':0,'skus':set()})
        if r.category=='Device': d['device']+=r.nett_amount
        elif r.category=='Macbook': d['macbook']+=r.nett_amount
        elif r.category=='ACC': d['acc']+=r.nett_amount
        if r.sku_key and r.category=='Device': d['skus'].add(r.sku_key)
    out=[]
    for (s,c,n),d in agg.items():
        total=d['device']+d['macbook']+d['acc']
        out.append({'salesman':s,'code':c,'name':n,'device':d['device'],'macbook':d['macbook'],'acc':d['acc'],'total':total,'bo':d['device']+d['macbook']>0,'qvo':total>=QVO_THRESHOLD,'sku':len(d['skus'])})
    return jsonify(out)

if __name__ == '__main__':
    with app.app_context(): db.create_all()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=os.environ.get('FLASK_DEBUG')=='1')
