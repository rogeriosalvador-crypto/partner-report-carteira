import os
import io
import json
from pathlib import Path
from functools import wraps
from flask import (
    Flask, render_template_string, request, redirect,
    url_for, session, send_file, jsonify, abort
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

BASE_DIR = Path(__file__).parent.absolute()
DATA_FILE = BASE_DIR / "data" / "relatorio.json"

GRUPOS = ["NAGUMO", "FESTVAL", "JACOMAR"]

PERIODOS_ORDER = ["D-1", "D-7", "D-15", "D-21", "MTD", "CONSOLIDADO"]

# ─── Credenciais por grupo ────────────────────────────────────────────────────
USERS = {
    "NAGUMO":  [("henrique","henrique"),("william","william"),("robert","robert"),("guilherme","guilherme")],
    "FESTVAL": [("anapaula","anapaula"),("vinicius","vinicius"),("vitoria","vitoria")],
    "JACOMAR": [("cristina","cristina"),("caio","caio")],
}

# ─── Dados default (fallback) ─────────────────────────────────────────────────
DEFAULT_DATA = {
    "data_referencia": "28/05/2026",
    "data_atualizacao": "28/05/2026 09:00",
    "periodos": {
        g: {
            p: {"gmv":0,"meta_gmv":0,"ating_pct":0,"pedidos":0,"aov":0,
                "cap_fat":0,"cancel_pct":0,"ruptura_pct":0,"er_pct":0,
                "inv_merchant":0,"pct_inv_merchant":0,"nps":0}
            for p in PERIODOS_ORDER
        } for g in GRUPOS
    },
    "lojas": {g: {p: [] for p in PERIODOS_ORDER} for g in GRUPOS}
}


def load_data():
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"Erro relatorio.json: {e}")
    return DEFAULT_DATA


def get_periodo(data, grupo, periodo):
    """Retorna métricas de um período seguro."""
    empty = {"gmv":0,"meta_gmv":0,"ating_pct":0,"pedidos":0,"aov":0,
             "cap_fat":0,"cancel_pct":0,"ruptura_pct":0,"er_pct":0,
             "inv_merchant":0,"pct_inv_merchant":0,"nps":0}
    return data.get("periodos",{}).get(grupo,{}).get(periodo, empty)


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "grupo" not in session:
            grupo = kwargs.get("grupo","").upper()
            return redirect(url_for("login_page", grupo=grupo.lower() or "nagumo"))
        return f(*args, **kwargs)
    return decorated


def fmt_brl(v):
    """Formata float como moeda BR."""
    try:
        return f"R$ {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return "R$ 0,00"


def fmt_pct(v, decimals=1):
    try:
        return f"{float(v):.{decimals}f}%".replace(".",",")
    except Exception:
        return "0%"


def semaforo_class(ating_pct):
    try:
        v = float(ating_pct)
        if v >= 100: return "verde"
        if v >= 85:  return "amarelo"
        return "vermelho"
    except Exception:
        return "vermelho"


# ════════════════════════════════════════════════════════════════
# TEMPLATES
# ════════════════════════════════════════════════════════════════

LOGIN_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Login — {{ grupo }}</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background: #0F0F1A; color: #f0f0f0;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      min-height: 100vh; display: flex; align-items: center; justify-content: center;
    }
    .login-box {
      background: #1A1A2E; border: 1px solid #2C3E50; border-radius: 16px;
      padding: 48px 40px; width: 100%; max-width: 400px;
      box-shadow: 0 8px 40px rgba(0,0,0,0.6);
    }
    .logo-row { display: flex; align-items: center; justify-content: center; gap: 10px; margin-bottom: 6px; }
    .logo-icon {
      width: 42px; height: 42px; background: #EA1D2C; border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      font-size: 22px; font-weight: 900; color: #fff;
    }
    .logo-text { font-size: 24px; font-weight: 800; color: #EA1D2C; }
    .grupo-badge {
      text-align: center; font-size: 13px; color: #B0BEC5; margin-bottom: 32px;
      letter-spacing: 1px; text-transform: uppercase;
    }
    label { display: block; font-size: 12px; color: #B0BEC5; margin-bottom: 6px;
            text-transform: uppercase; letter-spacing: 0.5px; }
    input[type=text], input[type=password] {
      width: 100%; background: #0F0F1A; border: 1px solid #2C3E50;
      border-radius: 8px; padding: 12px 14px; color: #fff; font-size: 15px;
      margin-bottom: 18px; outline: none; transition: border .2s;
    }
    input:focus { border-color: #EA1D2C; }
    button[type=submit] {
      width: 100%; background: #EA1D2C; color: #fff; border: none;
      border-radius: 8px; padding: 14px; font-size: 16px; font-weight: 700;
      cursor: pointer; transition: background .2s;
    }
    button:hover { background: #B71C2B; }
    .error-msg {
      background: rgba(231,76,60,.15); border: 1px solid #E74C3C;
      color: #E74C3C; border-radius: 8px; padding: 10px 14px;
      font-size: 13px; margin-bottom: 16px;
    }
  </style>
</head>
<body>
<div class="login-box">
  <div class="logo-row">
    <div class="logo-icon">i</div>
    <div class="logo-text">iFood</div>
  </div>
  <div class="grupo-badge">Parceiro {{ grupo }}</div>
  {% if erro %}<div class="error-msg">{{ erro }}</div>{% endif %}
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

DASHBOARD_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard {{ grupo }} — iFood</title>
  <style>
    :root {
      --bg: #0F0F1A; --card: #1A1A2E; --red: #EA1D2C; --red-dk: #B71C2B;
      --green: #27AE60; --yellow: #F39C12; --alert: #E74C3C;
      --txt: #FFFFFF; --txt2: #B0BEC5; --border: #2C3E50;
      --th: #16213E; --hover: #22304a;
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: var(--bg); color: var(--txt);
           font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           font-size: 14px; min-height: 100vh; }
    .container { max-width: 1200px; margin: 0 auto; padding: 16px; }

    /* TOPBAR */
    .topbar {
      background: var(--card); border-bottom: 1px solid var(--border);
      padding: 12px 20px; display: flex; align-items: center;
      justify-content: space-between; flex-wrap: wrap; gap: 10px;
      position: sticky; top: 0; z-index: 100;
    }
    .topbar-left { display: flex; align-items: center; gap: 12px; }
    .logo { font-size: 22px; font-weight: 800; color: var(--red); }
    .grupo-tag {
      background: var(--red); color: #fff; padding: 4px 12px;
      border-radius: 20px; font-size: 12px; font-weight: 700; letter-spacing: 1px;
    }
    .topbar-right { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .badge-update {
      background: var(--red-dk); color: #fff; padding: 5px 12px;
      border-radius: 20px; font-size: 11px; font-weight: 600;
    }
    .btn { padding: 7px 14px; border-radius: 8px; border: none; cursor: pointer;
           font-size: 13px; font-weight: 600; text-decoration: none;
           display: inline-flex; align-items: center; gap: 5px; }
    .btn-outline { background: transparent; border: 1px solid var(--border);
                   color: var(--txt2); }
    .btn-outline:hover { border-color: var(--red); color: var(--red); }
    .btn-red { background: var(--red); color: #fff; }
    .btn-red:hover { background: var(--red-dk); }
    .btn-gray { background: #2C3E50; color: var(--txt2); }
    .btn-gray:hover { background: #34495E; color: #fff; }

    /* PERIODO PILLS */
    .period-bar {
      background: var(--card); border: 1px solid var(--border);
      border-radius: 12px; padding: 12px 16px; margin: 16px 0;
      display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
    }
    .period-label { font-size: 12px; color: var(--txt2); font-weight: 600;
                    text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px; }
    .pill {
      padding: 7px 16px; border-radius: 20px; font-size: 13px; font-weight: 600;
      border: 1px solid var(--border); background: transparent; color: var(--txt2);
      cursor: pointer; transition: all .15s;
    }
    .pill:hover { border-color: var(--red); color: var(--red); }
    .pill.active { background: var(--red); border-color: var(--red); color: #fff; }

    /* KPIS GRID */
    .kpis-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
      gap: 12px; margin-bottom: 20px;
    }
    .kpi {
      background: var(--card); border: 1px solid var(--border);
      border-radius: 10px; padding: 16px;
    }
    .kpi-label { font-size: 11px; color: var(--txt2); text-transform: uppercase;
                 letter-spacing: 0.5px; margin-bottom: 6px; }
    .kpi-value { font-size: 22px; font-weight: 800; color: var(--txt); }
    .kpi-value.big { font-size: 26px; color: var(--red); }
    .kpi-value.verde { color: var(--green); }
    .kpi-value.amarelo { color: var(--yellow); }
    .kpi-value.vermelho { color: var(--alert); }
    .kpi-sub { font-size: 11px; color: var(--txt2); margin-top: 4px; }

    /* SEMAFORO */
    .sema { width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-right: 6px; }
    .sema.verde  { background: var(--green); box-shadow: 0 0 6px var(--green); }
    .sema.amarelo{ background: var(--yellow); box-shadow: 0 0 6px var(--yellow); }
    .sema.vermelho{ background: var(--alert); box-shadow: 0 0 6px var(--alert); }

    /* SECTION */
    .section { background: var(--card); border: 1px solid var(--border);
               border-radius: 12px; padding: 20px; margin-bottom: 20px; }
    .section-title { font-size: 16px; font-weight: 700; color: var(--red);
                     margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }

    /* TABLE */
    .tbl-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; }
    thead { background: var(--th); }
    th { text-align: left; padding: 10px 12px; font-size: 11px;
         text-transform: uppercase; letter-spacing: 0.5px; color: var(--txt2);
         font-weight: 600; white-space: nowrap; }
    td { padding: 11px 12px; border-top: 1px solid var(--border);
         font-size: 13px; white-space: nowrap; }
    tbody tr:hover { background: var(--hover); }
    .val-pos { color: var(--green); font-weight: 600; }
    .val-neg { color: var(--alert); font-weight: 600; }
    .val-warn { color: var(--yellow); font-weight: 600; }

    /* ATING BADGE inline */
    .ating-badge {
      display: inline-block; padding: 3px 8px; border-radius: 6px;
      font-weight: 700; font-size: 13px;
    }
    .ating-badge.verde  { background: var(--green); color: #fff; }
    .ating-badge.amarelo{ background: var(--yellow); color: #fff; }
    .ating-badge.vermelho{ background: var(--alert); color: #fff; }

    @media(max-width: 768px) {
      .kpis-grid { grid-template-columns: repeat(2, 1fr); }
      .topbar-right { flex-direction: column; align-items: flex-start; }
    }
    @media(max-width: 480px) {
      .kpis-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>

<!-- TOPBAR -->
<div class="topbar">
  <div class="topbar-left">
    <div class="logo">iFood</div>
    <div class="grupo-tag">{{ grupo }}</div>
  </div>
  <div class="topbar-right">
    <span class="badge-update">📅 {{ data_atualizacao }}</span>
    <a href="/download/excel/{{ grupo|lower }}" class="btn btn-gray">⬇ Excel</a>
    <button onclick="window.print()" class="btn btn-outline">🖨 Imprimir</button>
    <form method="POST" action="/logout" style="display:inline;">
      <button type="submit" class="btn btn-outline">Sair</button>
    </form>
  </div>
</div>

<div class="container">
  <!-- PERIODO SELECTOR -->
  <div class="period-bar">
    <span class="period-label">Período:</span>
    {% for p in periodos_order %}
    <button class="pill {% if p == periodo_ativo %}active{% endif %}"
            onclick="setPeriodo('{{ p }}')">{{ p }}</button>
    {% endfor %}
  </div>

  <!-- KPIs DO PERÍODO -->
  <div class="kpis-grid">
    <div class="kpi">
      <div class="kpi-label">GMV</div>
      <div class="kpi-value big">{{ fmt_brl(kpi.gmv) }}</div>
      <div class="kpi-sub">Meta: {{ fmt_brl(kpi.meta_gmv) }}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Atingimento</div>
      <div class="kpi-value {{ semaforo_class(kpi.ating_pct) }}">
        <span class="sema {{ semaforo_class(kpi.ating_pct) }}"></span>{{ fmt_pct(kpi.ating_pct) }}
      </div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Pedidos</div>
      <div class="kpi-value">{{ "{:,}".format(kpi.get("pedidos", 0)|int).replace(",",".") }}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">AOV</div>
      <div class="kpi-value">{{ fmt_brl(kpi.aov) }}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">CAP/FAT</div>
      <div class="kpi-value">{{ fmt_pct(kpi.cap_fat) }}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Cancelamento</div>
      <div class="kpi-value {% if kpi.cancel_pct|float > 5 %}val-neg{% endif %}">
        {{ fmt_pct(kpi.cancel_pct) }}
      </div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Ruptura</div>
      <div class="kpi-value {% if kpi.ruptura_pct|float > 1 %}val-neg{% endif %}">
        {{ fmt_pct(kpi.ruptura_pct) }}
      </div>
    </div>
    <div class="kpi">
      <div class="kpi-label">ER</div>
      <div class="kpi-value">{{ fmt_pct(kpi.er_pct) }}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">NPS</div>
      <div class="kpi-value {% if kpi.nps|float >= 80 %}verde{% elif kpi.nps|float >= 60 %}amarelo{% else %}vermelho{% endif %}">
        {{ fmt_pct(kpi.nps, 1)|replace('%','') }}
      </div>
    </div>
    {% if kpi.inv_merchant > 0 %}
    <div class="kpi">
      <div class="kpi-label">Inv. Merchant</div>
      <div class="kpi-value">{{ fmt_brl(kpi.inv_merchant) }}</div>
      <div class="kpi-sub">{{ fmt_pct(kpi.pct_inv_merchant) }} do GMV</div>
    </div>
    {% endif %}
  </div>

  <!-- GRÁFICOS SECTION -->
  <div class="section">
    <div class="section-title">📊 Evolução GMV por Período</div>
    <div id="chart-gmv"></div>
  </div>
  <div class="section">
    <div class="section-title">🎯 Atingimento% por Período</div>
    <div id="chart-ating"></div>
  </div>
  <div class="section">
    <div class="section-title">🏆 Top Lojas por GMV (Período Ativo)</div>
    <div id="chart-top-lojas"></div>
  </div>

  <!-- TABELA DE LOJAS (dinâmica via JS) -->
  <div class="section">
    <div class="section-title">🏪 Lojas — {{ grupo }} (<span id="periodo-lojas-label">{{ periodo_ativo }}</span>)</div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Loja</th>
            <th style="text-align:right">GMV</th>
            <th style="text-align:right">Pedidos</th>
            <th style="text-align:right">Ating%</th>
            <th style="text-align:right">Cancel%</th>
            <th style="text-align:right">Ruptura%</th>
            <th style="text-align:right">NPS</th>
          </tr>
        </thead>
        <tbody id="tabela-lojas-body">
          <tr><td colspan="7" style="text-align:center;color:#666;padding:20px">Carregando lojas...</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <!-- TABELA COMPARATIVO DE PERÍODOS -->
  <div class="section">
    <div class="section-title">📊 Comparativo de Períodos</div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Período</th>
            <th style="text-align:right">GMV</th>
            <th style="text-align:right">Meta GMV</th>
            <th style="text-align:right">Ating%</th>
            <th style="text-align:right">Pedidos</th>
            <th style="text-align:right">AOV</th>
            <th style="text-align:right">Cancel%</th>
            <th style="text-align:right">NPS</th>
          </tr>
        </thead>
        <tbody>
          {% for p in periodos_order %}
          {% set pr = todos_periodos[p] %}
          <tr {% if p == periodo_ativo %}style="background:rgba(234,29,44,.08);"{% endif %}>
            <td><strong>{{ p }}</strong></td>
            <td style="text-align:right">{{ fmt_brl(pr.gmv) }}</td>
            <td style="text-align:right">{{ fmt_brl(pr.meta_gmv) }}</td>
            <td style="text-align:right">
              <span class="ating-badge {{ semaforo_class(pr.ating_pct) }}">
                {{ fmt_pct(pr.ating_pct) }}
              </span>
            </td>
            <td style="text-align:right">{{ "{:,}".format(pr.get("pedidos", 0)|int).replace(",",".") }}</td>
            <td style="text-align:right">{{ fmt_brl(pr.aov) }}</td>
            <td style="text-align:right"
                class="{% if pr.cancel_pct > 5 %}val-neg{% endif %}">
              {{ fmt_pct(pr.cancel_pct) }}
            </td>
            <td style="text-align:right">{{ pr.nps }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>
</div>

<script>
// ─── Dados injetados pelo servidor ──────────────────────────────────────────
const TODOS_PERIODOS = {{ todos_periodos | tojson }};
const LOJAS_DICT     = {{ lojas_dict | tojson }};
let periodoAtivo     = "{{ periodo_ativo }}";
const ORDEM_PERIODOS = {{ periodos_order | tojson }};

// ─── Utilitários ────────────────────────────────────────────────────────────
function fmtBRL(v) {
  v = parseFloat(v) || 0;
  return "R$ " + v.toLocaleString("pt-BR", {minimumFractionDigits:2, maximumFractionDigits:2});
}
function semaforo(ating) {
  if (ating >= 100) return "ating-badge verde";
  if (ating >= 85)  return "ating-badge amarelo";
  return "ating-badge vermelho";
}

// ─── Troca de período: atualiza cards + label + lojas + gráficos ─────────────
function setPeriodo(p) {
  periodoAtivo = p;
  // Atualizar pills
  document.querySelectorAll(".pill").forEach(btn => {
    btn.classList.toggle("active", btn.textContent.trim() === p);
  });

  // Atualizar label de período na tabela de lojas
  const lbl = document.getElementById("periodo-lojas-label");
  if (lbl) lbl.textContent = p;

  // Atualizar KPI cards via URL para server-render (reload)
  // Manter compatibilidade: recarrega a página com o período selecionado
  const url = new URL(window.location.href);
  url.searchParams.set("periodo", p);

  // Atualizar lojas dinamicamente (sem reload)
  renderTabelaLojas(LOJAS_DICT[p] || []);
  renderChartTopLojas(LOJAS_DICT[p] || []);

  // Atualizar URL sem recarregar a página
  window.history.replaceState({}, "", url.toString());
}

// ─── Tabela de Lojas ────────────────────────────────────────────────────────
function renderTabelaLojas(lojas) {
  const tbody = document.getElementById("tabela-lojas-body");
  if (!tbody) return;
  if (!lojas || lojas.length === 0) {
    tbody.innerHTML = "<tr><td colspan=\"7\" style=\"text-align:center;color:#666;padding:20px\">Sem dados para este período</td></tr>";
    return;
  }
  tbody.innerHTML = lojas.map(l => {
    const ating = parseFloat(l.ating_pct) || 0;
    const cancelCls = (parseFloat(l.cancel_pct) || 0) > 5 ? "val-neg" : "";
    const ruptCls   = (parseFloat(l.ruptura_pct) || 0) > 1 ? "val-neg" : "";
    const nps       = parseFloat(l.nps) || 0;
    const npsCls    = nps >= 80 ? "val-pos" : nps >= 60 ? "val-warn" : "val-neg";
    return `<tr>
      <td><strong>${l.nome}</strong></td>
      <td style="text-align:right">${fmtBRL(l.gmv)}</td>
      <td style="text-align:right">${l.pedidos || 0}</td>
      <td style="text-align:right"><span class="${semaforo(ating)}">${ating.toFixed(1)}%</span></td>
      <td style="text-align:right" class="${cancelCls}">${(parseFloat(l.cancel_pct)||0).toFixed(1)}%</td>
      <td style="text-align:right" class="${ruptCls}">${(parseFloat(l.ruptura_pct)||0).toFixed(1)}%</td>
      <td style="text-align:right" class="${npsCls}">${nps.toFixed(0)}</td>
    </tr>`;
  }).join("");
}

// ─── Gráfico 1: GMV por período (barras horizontais CSS) ─────────────────────
function renderChartGMV() {
  const el = document.getElementById("chart-gmv");
  if (!el) return;
  const max = Math.max(...ORDEM_PERIODOS.map(p => parseFloat((TODOS_PERIODOS[p]||{}).gmv)||0), 1);
  el.innerHTML = ORDEM_PERIODOS.map(p => {
    const val = parseFloat((TODOS_PERIODOS[p]||{}).gmv)||0;
    const pct = (val / max * 100).toFixed(1);
    const isAtivo = p === periodoAtivo;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isAtivo?'#fff':'#B0BEC5'};font-weight:${isAtivo?'700':'400'}">${p}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;position:relative;overflow:hidden">
        <div style="width:${pct}%;background:${isAtivo?'#EA1D2C':'#4a4a6a'};height:100%;border-radius:4px;transition:width 0.3s"></div>
      </div>
      <span style="width:130px;text-align:right;font-size:12px;color:#FFF">${fmtBRL(val)}</span>
    </div>`;
  }).join("");
}

// ─── Gráfico 2: Atingimento% por período ─────────────────────────────────────
function renderChartAting() {
  const el = document.getElementById("chart-ating");
  if (!el) return;
  el.innerHTML = ORDEM_PERIODOS.map(p => {
    const val = parseFloat((TODOS_PERIODOS[p]||{}).ating_pct)||0;
    const pct = Math.min(val, 150);  // cap em 150% para display
    const cor = val >= 100 ? "#27AE60" : val >= 85 ? "#F39C12" : "#E74C3C";
    const isAtivo = p === periodoAtivo;
    return `<div style="display:flex;align-items:center;margin:6px 0;gap:10px">
      <span style="width:110px;font-size:12px;color:${isAtivo?'#fff':'#B0BEC5'};font-weight:${isAtivo?'700':'400'}">${p}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:24px;position:relative;overflow:hidden">
        <div style="width:${Math.min(pct/150*100,100).toFixed(1)}%;background:${isAtivo?cor:cor+'99'};height:100%;border-radius:4px;transition:width 0.3s"></div>
        <div style="position:absolute;left:0;top:0;height:100%;width:${(100/150*100).toFixed(1)}%;border-right:2px dashed #555"></div>
      </div>
      <span style="width:80px;text-align:right;font-size:12px;color:${cor};font-weight:700">${val.toFixed(1)}%</span>
    </div>`;
  }).join("");
}

// ─── Gráfico 3: Top lojas por GMV (período ativo) ─────────────────────────────
function renderChartTopLojas(lojas) {
  const el = document.getElementById("chart-top-lojas");
  if (!el) return;
  if (!lojas || lojas.length === 0) {
    el.innerHTML = "<p style=\"color:#666;text-align:center;padding:10px\">Sem dados para este período</p>";
    return;
  }
  const top10 = [...lojas].sort((a,b) => (parseFloat(b.gmv)||0) - (parseFloat(a.gmv)||0)).slice(0, 10);
  const max = Math.max(...top10.map(l => parseFloat(l.gmv)||0), 1);
  el.innerHTML = top10.map(l => {
    const val = parseFloat(l.gmv)||0;
    const pct = (val/max*100).toFixed(1);
    const ating = parseFloat(l.ating_pct)||0;
    const cor = ating >= 100 ? "#27AE60" : ating >= 85 ? "#F39C12" : "#E74C3C";
    return `<div style="display:flex;align-items:center;margin:5px 0;gap:10px">
      <span style="width:200px;font-size:11px;color:#B0BEC5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${l.nome}">${l.nome}</span>
      <div style="flex:1;background:#1A1A2E;border-radius:4px;height:22px;position:relative;overflow:hidden">
        <div style="width:${pct}%;background:${cor};height:100%;border-radius:4px;transition:width 0.3s;opacity:0.85"></div>
      </div>
      <span style="width:130px;text-align:right;font-size:11px;color:#FFF">${fmtBRL(val)}</span>
    </div>`;
  }).join("");
}

// ─── Inicialização ───────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", function() {
  renderTabelaLojas(LOJAS_DICT[periodoAtivo] || []);
  renderChartGMV();
  renderChartAting();
  renderChartTopLojas(LOJAS_DICT[periodoAtivo] || []);
});
</script>
</body>
</html>"""


# ════════════════════════════════════════════════════════════════
# ROTAS
# ════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return redirect(url_for("login_page", grupo="nagumo"))


@app.route("/<grupo>")
def login_page(grupo):
    g = grupo.upper()
    if g not in GRUPOS:
        abort(404)
    if session.get("grupo") == g:
        return redirect(url_for("dashboard", grupo=g.lower()))
    return render_template_string(LOGIN_TEMPLATE, grupo=g, erro=None)


@app.route("/login", methods=["POST"])
def login():
    grupo  = request.form.get("grupo","").upper()
    usuario = request.form.get("usuario","").strip().lower()
    senha   = request.form.get("senha","").strip()

    if grupo not in GRUPOS:
        abort(404)

    for u, s in USERS.get(grupo, []):
        if u == usuario and s == senha:
            session["grupo"]   = grupo
            session["usuario"] = usuario
            return redirect(url_for("dashboard", grupo=grupo.lower()))

    return render_template_string(
        LOGIN_TEMPLATE, grupo=grupo, erro="Usuário ou senha incorretos."
    )


@app.route("/dashboard/<grupo>")
@login_required
def dashboard(grupo):
    g = grupo.upper()
    if g not in GRUPOS:
        abort(404)
    if session.get("grupo") != g:
        return redirect(url_for("login_page", grupo=grupo))

    data = load_data()
    periodo_ativo = request.args.get("periodo", "D-1").upper()
    if periodo_ativo not in PERIODOS_ORDER:
        periodo_ativo = "D-1"

    todos_periodos = {p: get_periodo(data, g, p) for p in PERIODOS_ORDER}
    kpi   = todos_periodos[periodo_ativo]

    # lojas: suporte ao formato novo (dict por periodo) e ao formato legado (lista)
    raw_lojas = data.get("lojas", {}).get(g, {})
    if isinstance(raw_lojas, list):
        # formato legado: lista plana -> disponível apenas em D-1
        lojas_dict = {p: (raw_lojas if p == "D-1" else []) for p in PERIODOS_ORDER}
    else:
        lojas_dict = {p: raw_lojas.get(p, []) for p in PERIODOS_ORDER}

    return render_template_string(
        DASHBOARD_TEMPLATE,
        grupo=g,
        usuario=session.get("usuario"),
        data_referencia=data.get("data_referencia",""),
        data_atualizacao=data.get("data_atualizacao",""),
        periodo_ativo=periodo_ativo,
        periodos_order=PERIODOS_ORDER,
        kpi=kpi,
        todos_periodos=todos_periodos,
        lojas_dict=lojas_dict,
        fmt_brl=fmt_brl,
        fmt_pct=fmt_pct,
        semaforo_class=semaforo_class,
    )


@app.route("/logout", methods=["POST"])
def logout():
    grupo = session.get("grupo","NAGUMO").lower()
    session.clear()
    return redirect(url_for("login_page", grupo=grupo))


@app.route("/download/excel/<grupo>")
@login_required
def download_excel(grupo):
    g = grupo.upper()
    if g not in GRUPOS:
        abort(404)
    if session.get("grupo") != g:
        abort(403)

    data   = load_data()
    wb     = Workbook()
    RED_HEX   = "EA1D2C"
    DARK_HEX  = "16213E"
    HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
    RED_FILL  = PatternFill("solid", fgColor=RED_HEX)
    DARK_FILL = PatternFill("solid", fgColor=DARK_HEX)
    CENTER    = Alignment(horizontal="center", vertical="center")
    WRAP      = Alignment(wrap_text=True)

    def hdr(ws, row, cols):
        for col, val in enumerate(cols, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill = RED_FILL
            c.font = HEADER_FONT
            c.alignment = CENTER

    def sub_hdr(ws, row, cols):
        for col, val in enumerate(cols, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill = DARK_FILL
            c.font = Font(bold=True, color="B0BEC5", size=10)
            c.alignment = CENTER

    def brl(v):
        try: return float(v)
        except Exception: return 0.0

    # ── Aba 1: Comparativo de Períodos ──────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Comparativo"
    hdr(ws1, 1, ["Período","GMV","Meta GMV","Ating%","Pedidos","AOV",
                 "CAP/FAT%","Cancel%","Ruptura%","ER%","NPS","Inv.Merchant","% Inv"])
    for i, p in enumerate(PERIODOS_ORDER, 2):
        pr = get_periodo(data, g, p)
        row = [p, brl(pr["gmv"]), brl(pr["meta_gmv"]), brl(pr["ating_pct"]),
               int(pr.get("pedidos",0)), brl(pr.get("aov",0)),
               brl(pr.get("cap_fat",0)), brl(pr.get("cancel_pct",0)),
               brl(pr.get("ruptura_pct",0)), brl(pr.get("er_pct",0)),
               brl(pr.get("nps",0)), brl(pr.get("inv_merchant",0)),
               brl(pr.get("pct_inv_merchant",0))]
        for col, val in enumerate(row, 1):
            ws1.cell(row=i, column=col, value=val)
    for col in range(1, 14):
        ws1.column_dimensions[get_column_letter(col)].width = 14

    # ── Aba 2: Por Loja ──────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Por Loja")
    hdr(ws2, 1, ["Loja","GMV","Pedidos","Ating%","Cancel%","Ruptura%","NPS"])
    raw_lojas_xl = data.get("lojas",{}).get(g,{})
    if isinstance(raw_lojas_xl, list):
        lojas = raw_lojas_xl
    else:
        lojas = raw_lojas_xl.get("D-1", [])
    for i, loja in enumerate(lojas, 2):
        row = [loja.get("nome",""), brl(loja.get("gmv",0)),
               int(loja.get("pedidos",0)), brl(loja.get("ating_pct",0)),
               brl(loja.get("cancel_pct",0)), brl(loja.get("ruptura_pct",0)),
               brl(loja.get("nps",0))]
        for col, val in enumerate(row, 1):
            ws2.cell(row=i, column=col, value=val)
    ws2.column_dimensions["A"].width = 30
    for col in "BCDEFG":
        ws2.column_dimensions[col].width = 14

    # ── Aba 3: MTD ──────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("MTD")
    mtd = get_periodo(data, g, "MTD")
    hdr(ws3, 1, ["Métrica","Valor"])
    metricas = [
        ("GMV MTD", brl(mtd.get("gmv",0))),
        ("Meta GMV MTD", brl(mtd.get("meta_gmv",0))),
        ("Atingimento%", brl(mtd.get("ating_pct",0))),
        ("Pedidos", int(mtd.get("pedidos",0))),
        ("AOV", brl(mtd.get("aov",0))),
        ("CAP/FAT%", brl(mtd.get("cap_fat",0))),
        ("Cancelamento%", brl(mtd.get("cancel_pct",0))),
        ("Ruptura%", brl(mtd.get("ruptura_pct",0))),
        ("ER%", brl(mtd.get("er_pct",0))),
        ("NPS", brl(mtd.get("nps",0))),
        ("Inv. Merchant", brl(mtd.get("inv_merchant",0))),
        ("% Inv. Merchant", brl(mtd.get("pct_inv_merchant",0))),
    ]
    for i, (k, v) in enumerate(metricas, 2):
        ws3.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws3.cell(row=i, column=2, value=v)
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 18

    # Salva em buffer
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"relatorio_{g.lower()}_{data.get('data_referencia','').replace('/','')}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/atualizar", methods=["POST"])
def atualizar():
    """Recebe JSON completo e salva. Chamado pelo script interno, sem autenticação."""
    try:
        novo = request.get_json(force=True)
        if not novo:
            return jsonify({"error": "JSON inválido"}), 400
        DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
        DATA_FILE.write_text(json.dumps(novo, ensure_ascii=False, indent=2), encoding="utf-8")
        return jsonify({"ok": True, "data_atualizacao": novo.get("data_atualizacao","")})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)), debug=False)
