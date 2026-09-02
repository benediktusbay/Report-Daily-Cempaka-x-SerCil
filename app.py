#!/usr/bin/env python3
from flask import Flask, render_template_string, Markup
import locale

app = Flask(__name__)

# optional: set locale for number formatting (may depend on environment)
# We'll implement a simple formatter that uses dot as thousand separator.
def format_id(n):
    try:
        n = int(n)
    except (ValueError, TypeError):
        return "-"
    s = f"{n:,}".replace(",", ".")
    return s

app.jinja_env.filters['fmt'] = format_id

TEMPLATE = """
<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Dashboard</title>
  <style>
  :root{
    --primary-blue: #007bff;
    --muted-bg: #f6f7fb;
    --card-bg: #ffffff;
    --text: #222;
    --gap: 16px;
    --max-width:1200px;
  }
  *{box-sizing:border-box}
  body{font-family:Inter, system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial; color:var(--text); margin:0; background:#fafafa}
  .container{max-width:var(--max-width); margin:0 auto; padding:20px}
  /* Top nav */
  .nav{background:white; border-bottom:1px solid #eee; padding:10px 20px; border-radius:8px}
  .nav ul{list-style:none; margin:0; padding:0; display:flex; gap:12px; align-items:center; flex-wrap:wrap}
  .nav a{color:#333; text-decoration:none; padding:6px 8px}
  .nav a.active{font-weight:600; color:var(--primary-blue)}

  /* Cards row */
  .cards{display:flex; gap:var(--gap); margin-top:18px; flex-wrap:wrap}
  .card{background:var(--card-bg); border-radius:8px; padding:14px; box-shadow:0 1px 3px rgba(0,0,0,0.04); min-width:220px; flex:1 1 220px}
  .card-body{display:flex; flex-direction:column; gap:6px}
  .achievement-click{cursor:pointer; outline:none}
  .achievement-number{font-size:1.3rem; font-weight:700}
  .achievement-percent{color:#666}
  .progress{background:#e9ecef; height:10px; border-radius:6px; overflow:hidden; margin-top:8px}
  .progress-bar{background-color:var(--primary-blue); height:100%; transition:width .3s ease}

  /* Tables */
  .tables{margin-top:22px}
  .table-wrap{overflow:hidden}
  table.dashboard-table{width:100%; table-layout:fixed; border-collapse:collapse; background:white; border-radius:8px; overflow:hidden}
  table.dashboard-table th, table.dashboard-table td{padding:10px 12px; border-bottom:1px solid #f0f0f0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
  table.dashboard-table thead th{background:#fafafa; text-align:left; font-size:0.95rem; color:#444}
  /* Salesman kolom sama di kedua tabel */
  .col-salesman{width:24%; min-width:160px}
  .col-achievement{width:30%; min-width:140px}
  .col-percent{width:14%; min-width:90px}
  .achievement-cell{text-align:left; font-variant-numeric: tabular-nums;}

  /* floating input bottom (visual only) */
  .floating-input{position:fixed; left:50%; transform:translateX(-50%); bottom:18px; z-index:999; width:calc(100% - 48px); max-width:960px}
  .floating-box{background:white; border-radius:999px; padding:10px 14px; box-shadow:0 6px 24px rgba(0,0,0,0.08); display:flex; align-items:center; gap:10px}
  .floating-box input{flex:1; border:0; outline:none; font-size:15px}

  /* Modal */
  .modal{display:none; position:fixed; inset:0; z-index:2000; align-items:center; justify-content:center}
  .modal.open{display:flex}
  .modal-backdrop{position:absolute; inset:0; background:rgba(0,0,0,0.35)}
  .modal-panel{position:relative; background:#fff; border-radius:8px; max-width:480px; width:92%; z-index:2001; box-shadow:0 6px 24px rgba(0,0,0,0.18)}
  .modal-header{display:flex; align-items:center; justify-content:space-between; padding:12px 16px; border-bottom:1px solid #eee}
  .modal-body{padding:12px 16px}
  .modal-footer{padding:12px 16px; border-top:1px solid #eee; text-align:right}
  .modal-close{background:none; border:0; font-size:20px; cursor:pointer}

  /* small responsive tweaks */
  @media (max-width:900px){
    .container{padding:12px}
    .col-salesman{min-width:120px}
    .col-achievement{min-width:110px}
    table.dashboard-table th, table.dashboard-table td{padding:8px; font-size:13px}
    .cards{gap:10px}
    .floating-input{width:calc(100% - 28px)}
  }
  </style>
</head>
<body>
  <div class="container">
    <nav class="nav" aria-label="Main navigation">
      <ul>
        <li style="display:inline"><a href="#" class="active">Dashboard</a></li>
        <li style="display:inline"><a href="#">Stock</a></li>
        <li style="display:inline"><a href="#">Program</a></li>
        <li style="display:inline"><a href="#">Pricelist</a></li>
        <li style="display:inline"><a href="#">Incentive</a></li>
        <li style="display:inline"><a href="#">Target Master</a></li>
        <li style="display:inline"><a href="#">Users</a></li>
      </ul>
    </nav>

    <!-- Cards -->
    <section class="cards" aria-label="Summary cards">
      <!-- BO card (label kecil dihapus sesuai permintaan) -->
      <article class="card" aria-labelledby="bo-number">
        <div class="card-body">
          <div class="achievement-click" role="button" tabindex="0"
               data-target="{{ bo.target }}" data-achievement="{{ bo.achievement }}"
               data-percent="{{ bo.percent }}" data-gap="{{ bo.gap }}">
            <div id="bo-number" class="achievement-number">{{ bo.achievement|fmt }}</div>
            <div class="achievement-percent">{{ bo.percent }}%</div>
            <div class="progress" aria-hidden="true"><div class="progress-bar" style="width:{{ bo.percent }}%"></div></div>
          </div>
        </div>
      </article>

      <!-- QVO card (label kecil dihapus sesuai permintaan) -->
      <article class="card" aria-labelledby="qvo-number">
        <div class="card-body">
          <div class="achievement-click" role="button" tabindex="0"
               data-target="{{ qvo.target }}" data-achievement="{{ qvo.achievement }}"
               data-percent="{{ qvo.percent }}" data-gap="{{ qvo.gap }}">
            <div id="qvo-number" class="achievement-number">{{ qvo.achievement|fmt }}</div>
            <div class="achievement-percent">{{ qvo.percent }}%</div>
            <div class="progress" aria-hidden="true"><div class="progress-bar" style="width:{{ qvo.percent }}%"></div></div>
          </div>
        </div>
      </article>
    </section>

    <!-- Tables -->
    <section class="tables">
      <h3>Sales Achievement</h3>
      <div class="table-wrap" role="region" aria-label="Sales Achievement table">
        <table class="dashboard-table" role="table">
          <thead>
            <tr role="row">
              <th class="col-salesman" role="columnheader">Salesman</th>
              <th class="col-achievement" role="columnheader">Achievement</th>
              <th class="col-percent" role="columnheader">Achievement %</th>
            </tr>
          </thead>
          <tbody>
            {% for row in sales_achievements %}
            <tr role="row">
              <td class="col-salesman" role="cell">{{ row.salesman }}</td>
              <td class="achievement-cell col-achievement" role="cell">{{ row.achievement|fmt }}</td>
              <td class="col-percent" role="cell">{{ row.percent }}%</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>

      <h3 style="margin-top:18px">Distribution Quality</h3>
      <div class="table-wrap" role="region" aria-label="Distribution Quality table">
        <table class="dashboard-table" role="table">
          <thead>
            <tr>
              <th class="col-salesman">Salesman</th>
              <th>Quality</th>
            </tr>
          </thead>
          <tbody>
            {% for r in distribution_quality %}
            <tr>
              <td class="col-salesman">{{ r.salesman }}</td>
              <td>{{ r.quality }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </section>
  </div>

  <!-- Floating input UI (visual) -->
  <div class="floating-input" aria-hidden="true">
    <div class="floating-box">
      <span style="color:#888">Work with ChatGPT</span>
      <input type="text" placeholder="Ask for approval" aria-label="Work input" />
    </div>
  </div>

  <!-- Modal -->
  <div id="achievementModal" class="modal" aria-hidden="true" role="dialog" aria-labelledby="modalTitle" aria-modal="true">
    <div class="modal-backdrop" id="modalBackdrop"></div>
    <div class="modal-panel" role="document">
      <header class="modal-header">
        <h2 id="modalTitle">Detail Achievement</h2>
        <button id="modalClose" class="modal-close" aria-label="Tutup">&times;</button>
      </header>
      <section class="modal-body">
        <p><strong>Target:</strong> <span id="modal-target"></span></p>
        <p><strong>Achievement:</strong> <span id="modal-achievement"></span></p>
        <p><strong>Achievement %:</strong> <span id="modal-percent"></span></p>
        <p><strong>Gap to Target:</strong> <span id="modal-gap"></span></p>
      </section>
      <footer class="modal-footer">
        <button id="modalCloseBtn">Tutup</button>
      </footer>
    </div>
  </div>

  <script>
  // Simple modal + handlers
  document.addEventListener('DOMContentLoaded', function() {
    const modal = document.getElementById('achievementModal');
    const modalBackdrop = document.getElementById('modalBackdrop');
    const modalClose = document.getElementById('modalClose');
    const modalCloseBtn = document.getElementById('modalCloseBtn');

    function fmt(n){
      if(n === null || n === undefined) return '-';
      return n.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ".");
    }

    function openModal(data){
      document.getElementById('modal-target').textContent = fmt(data.target);
      document.getElementById('modal-achievement').textContent = fmt(data.achievement);
      document.getElementById('modal-percent').textContent = (data.percent !== undefined ? data.percent + '%' : '-');
      document.getElementById('modal-gap').textContent = fmt(data.gap);
      modal.classList.add('open');
      modal.setAttribute('aria-hidden','false');
      modalClose.focus();
    }

    function closeModal(){
      modal.classList.remove('open');
      modal.setAttribute('aria-hidden','true');
    }

    document.querySelectorAll('.achievement-click').forEach(function(el){
      el.addEventListener('click', function(){
        const data = {
          target: Number(el.dataset.target),
          achievement: Number(el.dataset.achievement),
          percent: el.dataset.percent,
          gap: Number(el.dataset.gap)
        };
        openModal(data);
      });
      el.addEventListener('keydown', function(e){
        if(e.key === 'Enter' || e.key === ' '){
          e.preventDefault();
          el.click();
        }
      });
    });

    modalBackdrop && modalBackdrop.addEventListener('click', closeModal);
    modalClose && modalClose.addEventListener('click', closeModal);
    modalCloseBtn && modalCloseBtn.addEventListener('click', closeModal);

    document.addEventListener('keydown', function(e){
      if(e.key === 'Escape' && modal.classList.contains('open')) closeModal();
    });
  });
  </script>
</body>
</html>
"""

# Sample data (replace with real data or integrate DB)
def sample_sales_achievements():
    return [
        {"salesman": "Andi Pratama", "achievement": 1500000, "percent": 75, "target": 2000000, "gap": 500000},
        {"salesman": "Budi Santoso", "achievement": 900000, "percent": 45, "target": 2000000, "gap": 1100000},
        {"salesman": "Citra", "achievement": 2000000, "percent": 100, "target": 2000000, "gap": 0},
    ]

def sample_distribution_quality():
    return [
        {"salesman": "Andi Pratama", "quality": 90},
        {"salesman": "Budi Santoso", "quality": 85},
        {"salesman": "Citra", "quality": 92},
    ]

@app.route("/")
def dashboard():
    # summary cards (BO & QVO) sample
    bo = {"target": 2000000, "achievement": 1500000, "percent": 75, "gap": 500000}
    qvo = {"target": 2000000, "achievement": 900000, "percent": 45, "gap": 1100000}
    return render_template_string(TEMPLATE,
                                  bo=bo,
                                  qvo=qvo,
                                  sales_achievements=sample_sales_achievements(),
                                  distribution_quality=sample_distribution_quality(),
                                  active='dashboard')

if __name__ == "__main__":
    # debug mode for local development
    app.run(debug=True)
