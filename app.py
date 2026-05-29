"""
Partner Report — Dashboard por grupo (login por parceiro).
Visual escuro iFood. Cada período mostra dados distintos de lojas.
"""
import os, io, json
from pathlib import Path
from functools import wraps
from flask import (
    Flask, render_template_string, request, redirect,
    url_for, session, send_file, jsonify
)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

BASE_DIR = Path(__file__).parent.absolute()
DATA_FILE = BASE_DIR / "data" / "relatorio.json"

GRUPOS = ["NAGUMO", "FESTVAL", "JACOMAR"]
PERIODOS_ORDER = ["D-1", "D-7", "D-15", "D-21", "MTD", "CONSOLIDADO"]

# ── Credenciais por grupo ────────────────────────────────────────────────────
USERS = {
    "NAGUMO":  [("henrique","henrique"),("william","william"),("robert","robert"),("guilherme","guilherme")],
    "FESTVAL": [("anapaula","anapaula"),("vinicius","vinicius"),("vitoria","vitoria")],
    "JACOMAR": [("cristina","cristina"),("caio","caio")],
}

# ── Default data ─────────────────────────────────────────────────────────────
DEFAULT_DATA = {
    "data_referencia": "27/05/2026",
    "data_atualizacao": "28/05/2026 00:00",
    "periodos": {
        g: {p: {"gmv":0,"meta_gmv":0,"ating_pct":0,"pedidos":0,"aov":0,"cancel_pct":0,"ruptura_pct":0,"nps":None}
            for p in PERIODOS_ORDER}
        for g in GRUPOS
    },
    "lojas": {g: {p: [] for p in PERIODOS_ORDER} for g in GRUPOS}
}

def load_data():
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"[WARN] relatorio.json: {e}")
    return DEFAULT_DATA

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "grupo" not in session:
            grupo = kwargs.get("grupo","").upper()
            return redirect(url_for("login_page", grupo=grupo.lower() or "nagumo"))
        return f(*args, **kwargs)
    return decorated

# ════════════════════════════════════════════════════════════════
# LOGIN TEMPLATE
# ════════════════════════════════════════════════════════════════
LOGIN_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Login — {{ grupo }}</title>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{background:#0F0F1A;color:#f0f0f0;
         font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
         min-height:100vh;display:flex;align-items:center;justify-content:center}
    .box{background:#1A1A2E;border:1px solid #2C3E50;border-radius:16px;
         padding:48px 40px;width:100%;max-width:400px;box-shadow:0 8px 40px rgba(0,0,0,.6)}
    .logo-row{display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:6px}
    .logo-icon{width:42px;height:42px;background:#EA1D2C;border-radius:10px;
               display:flex;align-items:center;justify-content:center;
               font-size:22px;font-weight:900;color:#fff}
    .logo-txt{font-size:24px;font-weight:800;color:#EA1D2C}
    .badge{text-align:center;font-size:13px;color:#B0BEC5;margin-bottom:32px;
           letter-spacing:1px;text-transform:uppercase}
    label{display:block;font-size:12px;color:#B0BEC5;margin-bottom:6px;
          text-transform:uppercase;letter-spacing:.5px}
    input[type=text],input[type=password]{width:100%;background:#0F0F1A;
      border:1px solid #2C3E50;border-radius:8px;padding:12px 14px;
      color:#fff;font-size:15px;margin-bottom:18px;outline:none;transition:border .2s}
    input:focus{border-color:#EA1D2C}
    button[type=submit]{width:100%;background:#EA1D2C;color:#fff;border:none;
      border-radius:8px;padding:14px;font-size:16px;font-weight:700;cursor:pointer}
    button:hover{background:#B71C2B}
    .err{background:rgba(231,76,60,.15);border:1px solid #E74C3C;color:#E74C3C;
         border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:16px}
  </style>
</head>
<body>
<div class="box">
  <div class="logo-row"><div class="logo-icon">i</div><div class="logo-txt">iFood</div></div>
  <div class="badge">Parceiro {{ grupo }}</div>
  {% if erro %}<div class="err">{{ erro }}</div>{% endif %}
  <form method="POST" action="/login">
    <input type="hidden" name="grupo" value="{{ grupo }}">
    <label>Usuário</label>
    <input type="text" name="usuario" autocomplete="username" required>
    <label>Senha</label>
    <input type="password" name="senha" autocomplete="current-password" required>
    <button type="submit">Entrar</button>
  </form>
</div>
</body>
</html>"""

# ════════════════════════════════════════════════════════════════
# DASHBOARD TEMPLATE
# ════════════════════════════════════════════════════════════════
DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard {{ grupo }} — iFood</title>
  <style>
    :root{--bg:#0F0F1A;--card:#1A1A2E;--red:#EA1D2C;--red-dk:#B71C2B;
          --green:#27AE60;--yellow:#F39C12;--alert:#E74C3C;
          --txt:#FFFFFF;--txt2:#B0BEC5;--border:#2C3E50;--th:#16213E;--hover:#22304a}
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{background:var(--bg);color:var(--txt);
         font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
         font-size:14px;min-height:100vh}
    .container{max-width:1200px;margin:0 auto;padding:16px}

    .topbar{background:var(--card);border-bottom:1px solid var(--border);
            padding:12px 20px;display:flex;align-items:center;
            justify-content:space-between;flex-wrap:wrap;gap:10px;
            position:sticky;top:0;z-index:100}
    .topbar-left{display:flex;align-items:center;gap:12px}
    .logo{font-size:22px;font-weight:800;color:var(--red)}
    .grupo-name{font-size:16px;font-weight:700;color:var(--txt)}
    .topbar-right{display:flex;gap:8px;flex-wrap:wrap}
    .btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;
         border:none;border-radius:8px;font-size:13px;font-weight:600;
         cursor:pointer;text-decoration:none;transition:opacity .15s}
    .btn-red{background:var(--red);color:#fff}
    .btn-out{background:transparent;color:var(--txt2);border:1px solid var(--border)}
    .btn:hover{opacity:.85}

    .update-info{background:var(--card);border-bottom:1px solid var(--border);
                 padding:8px 20px;font-size:12px;color:var(--txt2);
                 display:flex;gap:16px;flex-wrap:wrap}
    .update-info span{color:var(--txt)}

    .period-bar{display:flex;gap:6px;flex-wrap:wrap;padding:14px 20px 0}
    .period-pill{padding:6px 16px;border-radius:20px;border:1px solid var(--border);
                 background:var(--card);color:var(--txt2);
                 cursor:pointer;font-size:13px;font-weight:600;transition:all .15s}
    .period-pill.active{background:var(--red);color:#fff;border-color:var(--red)}

    .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
              gap:12px;padding:16px 20px}
    .kpi-card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:14px}
    .kpi-label{font-size:11px;color:var(--txt2);text-transform:uppercase;
               letter-spacing:.5px;margin-bottom:6px}
    .kpi-value{font-size:20px;font-weight:800;color:var(--txt)}
    .kpi-value.big{font-size:26px;color:var(--red)}
    .badge{display:inline-block;padding:3px 9px;border-radius:5px;
           font-weight:700;font-size:14px;margin-left:6px}
    .badge.verde{background:var(--green);color:#fff}
    .badge.amarelo{background:var(--yellow);color:#fff}
    .badge.vermelho{background:var(--alert);color:#fff}
    .ating-badge{display:inline-block;padding:2px 7px;border-radius:5px;font-weight:700;font-size:12px}
    .ating-badge.verde{background:var(--green);color:#fff}
    .ating-badge.amarelo{background:var(--yellow);color:#fff}
    .ating-badge.vermelho{background:var(--alert);color:#fff}

    .section{background:var(--card);border:1px solid var(--border);
             border-radius:12px;margin:0 20px 20px;padding:18px}
    .section-title{font-size:16px;font-weight:700;color:var(--red);margin-bottom:14px}
    .table-wrap{overflow-x:auto}
    table{width:100%;border-collapse:collapse}
    thead{background:var(--th)}
    th{text-align:left;padding:10px 12px;font-size:11px;text-transform:uppercase;
       letter-spacing:.5px;color:var(--txt2);font-weight:600;white-space:nowrap}
    td{padding:10px 12px;border-top:1px solid var(--border);font-size:13px}
    tbody tr:hover{background:var(--hover)}
    .val-pos{color:var(--green);font-weight:600}
    .val-neg{color:var(--alert);font-weight:600}
    .val-warn{color:var(--yellow);font-weight:600}
    .sem{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:6px}
    .sem-v{background:var(--green);box-shadow:0 0 6px var(--green)}
    .sem-a{background:var(--yellow);box-shadow:0 0 6px var(--yellow)}
    .sem-r{background:var(--alert);box-shadow:0 0 6px var(--alert)}

    @media(max-width:600px){
      .kpi-grid{grid-template-columns:repeat(2,1fr)}
      th,td{padding:8px 8px}
    }
  </style>
</head>
<body>

<div class="topbar">
  <div class="topbar-left">
    <div class="logo">iFood</div>
    <div class="grupo-name">{{ grupo }}</div>
  </div>
  <div class="topbar-right">
    <a class="btn btn-red" href="/download/excel/{{ grupo|lower }}" download>⬇ Excel</a>
    <button class="btn btn-out" onclick="window.print()">🖨 Imprimir</button>
    <a class="btn btn-out" href="/logout">Sair</a>
  </div>
</div>

<div class="update-info">
  <div>Referência: <span>{{ data_referencia }}</span></div>
  <div>Atualizado: <span>{{ data_atualizacao }}</span></div>
</div>

<!-- PERIOD PILLS -->
<div class="period-bar" id="periodBar"></div>

<!-- KPI CARDS -->
<div class="kpi-grid" id="kpiGrid"></div>

<!-- GMV Chart -->
<div class="section">
  <div class="section-title">📊 Evolução GMV por Período</div>
  <div id="chartGmv"></div>
</div>

<!-- Top Lojas Chart -->
<div class="section">
  <div class="section-title">🏆 Top Lojas por GMV — <span id="chartPeriodLabel"></span></div>
  <div id="chartTop"></div>
</div>

<!-- Stores Table -->
<div class="section">
  <div class="section-title">🏪 Lojas — <span id="tableLabel"></span></div>
  <div class="table-wrap">
    <table>
      <thead id="storesThead"></thead>
      <tbody id="storesTbody"></tbody>
    </table>
  </div>
</div>

<script>
// ── Data injected by server ──────────────────────────────────────────────────
const TODOS_PERIODOS = {{ todos_periodos | tojson }};  // {periodo: {gmv,pedidos,...}}
const LOJAS_DICT     = {{ lojas_dict | tojson }};      // {periodo: [{nome,gmv,...}]}
const PERIODOS_ORDER = {{ periodos_order | tojson }};
const PERIOD_LABELS  = {'D-1':'D-1','D-7':'D-7','D-15':'D-15','D-21':'D-21','MTD':'MTD','CONSOLIDADO':'Consolidado'};

let periodoAtivo = 'D-1';

// ── Formatters ───────────────────────────────────────────────────────────────
const fmtBRL = v => v==null?'R$ —':'R$ '+Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtPct = v => v==null?'—':Number(v).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%';
const fmtNum = v => v==null?'—':Number(v).toLocaleString('pt-BR');
const semCls  = a => a>=100?'sem-v':a>=85?'sem-a':'sem-r';
const badgeCls = a => a>=100?'verde':a>=85?'amarelo':'vermelho';

// ── Period bar ────────────────────────────────────────────────────────────────
function renderPeriodBar() {
  document.getElementById('periodBar').innerHTML = PERIODOS_ORDER.map(p =>
    `<div class="period-pill ${p===periodoAtivo?'active':''}" onclick="setPeriodo('${p}')">${PERIOD_LABELS[p]||p}</div>`
  ).join('');
}

// ── KPI Cards ─────────────────────────────────────────────────────────────────
function renderKPICards(kpi) {
  const ating = kpi.ating_pct||0;
  const isMTD = periodoAtivo==='MTD'||periodoAtivo==='CONSOLIDADO';
  const cards = [];

  cards.push(`<div class="kpi-card">
    <div class="kpi-label">GMV</div>
    <div class="kpi-value big">${fmtBRL(kpi.gmv)}
      ${kpi.ating_pct!=null?`<span class="badge ${badgeCls(ating)}">${fmtPct(ating)}</span>`:''}
    </div>
  </div>`);

  if (kpi.meta_gmv) {
    cards.push(`<div class="kpi-card">
      <div class="kpi-label">Meta GMV</div>
      <div class="kpi-value">${fmtBRL(kpi.meta_gmv)}</div>
    </div>`);
  }

  cards.push(`<div class="kpi-card">
    <div class="kpi-label">Pedidos</div>
    <div class="kpi-value">${fmtNum(kpi.pedidos)}</div>
  </div>`);

  if (!isMTD) {
    if (kpi.aov) cards.push(`<div class="kpi-card">
      <div class="kpi-label">AOV</div>
      <div class="kpi-value">${fmtBRL(kpi.aov)}</div>
    </div>`);

    if (kpi.cancel_pct!=null) {
      const cls = kpi.cancel_pct>7?'val-neg':kpi.cancel_pct>5?'val-warn':'';
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">Cancelamento</div>
        <div class="kpi-value ${cls}">${fmtPct(kpi.cancel_pct)}</div>
      </div>`);
    }

    if (kpi.ruptura_pct!=null) {
      const cls = kpi.ruptura_pct>3?'val-neg':kpi.ruptura_pct>1?'val-warn':'';
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">Ruptura</div>
        <div class="kpi-value ${cls}">${fmtPct(kpi.ruptura_pct)}</div>
      </div>`);
    }

    if (kpi.nps!=null) {
      const cls = kpi.nps>=70?'val-pos':kpi.nps>=50?'val-warn':'val-neg';
      cards.push(`<div class="kpi-card">
        <div class="kpi-label">NPS</div>
        <div class="kpi-value ${cls}">${Number(kpi.nps).toLocaleString('pt-BR',{minimumFractionDigits:1})}</div>
      </div>`);
    }
  }
  document.getElementById('kpiGrid').innerHTML = cards.join('');
}

// ── Tabela de lojas — usa LOJAS_DICT[periodo] ─────────────────────────────────
const PERIODO_REF = {
  'D-1':'D-7', 'D-7':'D-15', 'D-15':'D-21',
  'D-21':'MTD', 'MTD':null, 'CONSOLIDADO':'MTD'
};

function calcAtingDinamico(loja, periodo) {
  const ref = PERIODO_REF[periodo];
  if (!ref) return parseFloat(loja.ating_pct_meta) || 0;
  const lojas_ref = LOJAS_DICT[ref] || [];
  const loja_ref = lojas_ref.find(l => l.nome === loja.nome);
  if (!loja_ref || !loja_ref.gmv) return 0;
  return Math.round(loja.gmv / loja_ref.gmv * 1000) / 10;
}

function renderTabelaLojas(lojas, periodo) {
  const tbody = document.getElementById("storesTbody");
  const thead = document.getElementById("storesThead");
  if (!tbody) return;
  if (!lojas || lojas.length === 0) {
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = "<tr><td colspan='11' style='text-align:center;color:#666;padding:20px'>Sem dados para este período</td></tr>";
    return;
  }
  if (thead) thead.innerHTML = `<tr>
    <th>Loja</th>
    <th style="text-align:right">GMV</th>
    <th style="text-align:right">Pedidos</th>
    <th style="text-align:right">Ating%</th>
    <th style="text-align:right">Cancel%</th>
    <th style="text-align:right">Ruptura%</th>
    <th style="text-align:right">SLA Sep%</th>
    <th style="text-align:right">Online%</th>
    <th style="text-align:right">GMV Rupt</th>
    <th style="text-align:right">GMV Recup</th>
    <th style="text-align:right">NSU%</th>
  </tr>`;
  tbody.innerHTML = lojas.map(l => {
    const ating = calcAtingDinamico(l, periodo);
    const atBadge = ating >= 100 ? "ating-badge verde" : ating >= 85 ? "ating-badge amarelo" : "ating-badge vermelho";
    const fmt1 = v => v != null ? v.toFixed(1)+'%' : '—';
    const fmtR = v => v ? fmtBRL(v) : '—';
    const slaCls = (l.sla_sep_pct||0)>=85?"val-pos":(l.sla_sep_pct||0)>=70?"val-warn":"val-neg";
    const onlineCls = (l.online_pct||0)>=95?"val-pos":(l.online_pct||0)>=80?"val-warn":"val-neg";
    const cancelCls = (l.cancel_pct||0)>5?"val-neg":"";
    const ruptCls = (l.ruptura_pct||0)>1?"val-neg":"";
    const nsuCls = (l.nsu_pct||0)<=2?"val-pos":(l.nsu_pct||0)<=5?"val-warn":"val-neg";
    return `<tr>
      <td><strong>${l.nome}</strong></td>
      <td style="text-align:right">${fmtBRL(l.gmv)}</td>
      <td style="text-align:right">${l.pedidos||0}</td>
      <td style="text-align:right"><span class="${atBadge}">${ating.toFixed(1)}%</span></td>
      <td style="text-align:right" class="${cancelCls}">${fmt1(l.cancel_pct)}</td>
      <td style="text-align:right" class="${ruptCls}">${fmt1(l.ruptura_pct)}</td>
      <td style="text-align:right" class="${slaCls}">${fmt1(l.sla_sep_pct)}</td>
      <td style="text-align:right" class="${onlineCls}">${fmt1(l.online_pct)}</td>
      <td style="text-align:right">${fmtR(l.gmv_rupturado)}</td>
      <td style="text-align:right">${fmtR(l.gmv_recuperado)}</td>
      <td style="text-align:right" class="${nsuCls}">${fmt1(l.nsu_pct)}</td>
    </tr>`;
  }).join("");
}

function renderChartTopLojas(lojas) {
  const el = document.getElementById('chartTop');
  if (!lojas || !lojas.length) {
    el.innerHTML='<p style="color:#666;text-align:center;padding:10px">Sem dados para este período</p>';
    return;
  }
  const top10 = [...lojas].sort((a,b)=>(b.gmv||0)-(a.gmv||0)).slice(0,10);
  const max = Math.max(...top10.map(l=>l.gmv||0),1);
  el.innerHTML = top10.map(l => {
    const val = l.gmv||0;
    const pct = (val/max*100).toFixed(1);
    const ating = l.ating_pct||0;
    const cor = ating>=100?'#27AE60':ating>=85?'#F39C12':'#E74C3C';
    return `<div style="display:flex;align-items:center;margin:5px 0;gap:10px">
      <span style="width:200px;font-size:11px;color:#B0BEC5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${l.nome}</span>
      <div style="flex:1;background:#16213E;border-radius:4px;height:22px;overflow:hidden">
        <div style="width:${pct}%;background:${cor};height:100%;border-radius:4px;opacity:.85"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:11px;color:#FFF">R$ ${val.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
    </div>`;
  }).join('');
}

// ── GMV evolution chart ───────────────────────────────────────────────────────
function renderChartGmv() {
  const max = Math.max(...PERIODOS_ORDER.map(p=>parseFloat((TODOS_PERIODOS[p]||{}).gmv)||0),1);
  document.getElementById('chartGmv').innerHTML = PERIODOS_ORDER.map(p => {
    const val = parseFloat((TODOS_PERIODOS[p]||{}).gmv)||0;
    const pct = (val/max*100).toFixed(1);
    const isActive = p===periodoAtivo;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isActive?'#fff':'#B0BEC5'};font-weight:${isActive?700:400}">${PERIOD_LABELS[p]||p}</span>
      <div style="flex:1;background:#16213E;border-radius:4px;height:24px;overflow:hidden">
        <div style="width:${pct}%;background:${isActive?'#EA1D2C':'#4a4a6a'};height:100%;border-radius:4px;transition:width .3s"></div>
      </div>
      <span style="width:150px;text-align:right;font-size:12px;color:#FFF">R$ ${val.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>
    </div>`;
  }).join('');
}

// ── Main: set período ─────────────────────────────────────────────────────────
function setPeriodo(p) {
  periodoAtivo = p;
  const label = PERIOD_LABELS[p]||p;
  renderPeriodBar();
  renderKPICards(TODOS_PERIODOS[p]||{});
  renderTabelaLojas(LOJAS_DICT[p]||[], p);        // ← dados do período p (distintos!)
  renderChartTopLojas(LOJAS_DICT[p]||[]);
  renderChartGmv();
  document.getElementById('tableLabel').textContent = label;
  document.getElementById('chartPeriodLabel').textContent = label;
}

// Boot
setPeriodo('D-1');
</script>
</body>
</html>"""

# ════════════════════════════════════════════════════════════════
# EXCEL
# ════════════════════════════════════════════════════════════════
def build_excel_partner(data, grupo):
    wb = Workbook()
    HDR = PatternFill("solid", fgColor="EA1D2C")
    HFONT = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center")

    ws1 = wb.active; ws1.title = "Resumo Períodos"
    ws1.column_dimensions['A'].width = 16
    for c in 'BCDEFGH': ws1.column_dimensions[c].width = 18
    ws1.append(['Período','GMV','Meta GMV','Ating%','Pedidos','Cancel%','Ruptura%'])
    for cell in ws1[1]:
        cell.fill=HDR; cell.font=HFONT; cell.alignment=center

    for p in PERIODOS_ORDER:
        kpi = data.get('periodos',{}).get(grupo,{}).get(p,{})
        ws1.append([p, kpi.get('gmv',0), kpi.get('meta_gmv',0),
                    kpi.get('ating_pct',0), kpi.get('pedidos',0),
                    kpi.get('cancel_pct'), kpi.get('ruptura_pct')])

    # D-1 lojas sheet
    ws2 = wb.create_sheet("Lojas D-1")
    ws2.column_dimensions['A'].width = 40
    for c in 'BCDEF': ws2.column_dimensions[c].width = 16
    ws2.append(['Loja','GMV','Pedidos','Cancel%','Ruptura%','NPS'])
    for cell in ws2[1]:
        cell.fill=HDR; cell.font=HFONT; cell.alignment=center
    for l in data.get('lojas',{}).get(grupo,{}).get('D-1',[]):
        ws2.append([l.get('nome',''),l.get('gmv',0),l.get('pedidos',0),
                    l.get('cancel_pct'),l.get('ruptura_pct'),l.get('nps')])

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ════════════════════════════════════════════════════════════════
# ROUTES
# ════════════════════════════════════════════════════════════════
@app.route('/')
def home():
    return redirect(url_for('login_page', grupo='nagumo'))

@app.route('/login/<grupo>')
@app.route('/login')
def login_page(grupo='nagumo'):
    return render_template_string(LOGIN_HTML, grupo=grupo.upper(), erro=None)

@app.route('/login', methods=['POST'])
def login_post():
    grupo = request.form.get('grupo','').upper()
    usuario = request.form.get('usuario','').strip().lower()
    senha = request.form.get('senha','').strip()

    creds = USERS.get(grupo, [])
    for u, p in creds:
        if u == usuario and p == senha:
            session['grupo'] = grupo
            session['usuario'] = usuario
            return redirect(url_for('dashboard', grupo=grupo.lower()))

    return render_template_string(LOGIN_HTML, grupo=grupo, erro="Usuário ou senha inválidos")

@app.route('/logout')
def logout():
    grupo = session.get('grupo','nagumo').lower()
    session.clear()
    return redirect(url_for('login_page', grupo=grupo))

@app.route('/dashboard/<grupo>')
@login_required
def dashboard(grupo):
    grupo_up = grupo.upper()
    if session.get('grupo') != grupo_up:
        return redirect(url_for('dashboard', grupo=session['grupo'].lower()))

    data = load_data()

    # Build todos_periodos: {periodo: {kpis...}}
    todos_periodos = {}
    for p in PERIODOS_ORDER:
        todos_periodos[p] = data.get('periodos',{}).get(grupo_up,{}).get(p,{})

    # Build lojas_dict: {periodo: [lojas...]}
    lojas_dict = {}
    for p in PERIODOS_ORDER:
        lojas_dict[p] = data.get('lojas',{}).get(grupo_up,{}).get(p,[])

    return render_template_string(
        DASHBOARD_HTML,
        grupo=grupo_up,
        data_referencia=data.get('data_referencia',''),
        data_atualizacao=data.get('data_atualizacao',''),
        todos_periodos=todos_periodos,
        lojas_dict=lojas_dict,
        periodos_order=PERIODOS_ORDER,
    )

@app.route('/download/excel/<grupo>')
@login_required
def download_excel(grupo):
    if not OPENPYXL_OK:
        return jsonify({"error":"openpyxl não disponível"}), 500
    grupo_up = grupo.upper()
    if session.get('grupo') != grupo_up:
        return redirect(url_for('dashboard', grupo=session['grupo'].lower()))
    data = load_data()
    buf = build_excel_partner(data, grupo_up)
    ref = data.get('data_referencia','').replace('/','.')
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=f"relatorio_{grupo_up}_{ref}.xlsx")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8081))
    app.run(host='0.0.0.0', port=port, debug=False)
