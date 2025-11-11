from __future__ import annotations
from pathlib import Path
from datetime import datetime
import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for

# ========= 可調整區 =========
DATA_DIR = Path(__file__).parent / "data"

LOAD_XLSX = DATA_DIR / "load.xlsx"         # ← 你的「中西區_逐日逐時_LONG(MW).xlsx」可改名成這個，或改這行
PV_XLSX   = DATA_DIR / "pv.xlsx"           # ← 你的「自己推的中西區pv發電量.xlsx」可改名成這個，或改這行

PV_SHEET = 0 
# 欄位名稱（若你的欄位不同，改這裡即可）
LOAD_COLS = {"datetime": "datetime", "value": "load"}        # 里負載檔的欄位對應
PV_COLS   = {"datetime": "datetime", "value": "generator"}   # PV 檔的欄位對應
# ========= 可調整區 =========

app = Flask(__name__)

def _ensure_exists(p: Path):
    if not p.exists():
        raise FileNotFoundError(f"找不到資料檔：{p}")

def list_villages_from_load() -> list[str]:
    """列出負載檔的所有工作表（視為里名）。"""
    _ensure_exists(LOAD_XLSX)
    xls = pd.ExcelFile(LOAD_XLSX)
    return xls.sheet_names

def day_series_from_sheet(
    xlsx: Path,
    sheet: str | int | None,
    date_str: str,
    col_map: dict[str, str],
) -> pd.DataFrame:
    """
    從指定檔案/工作表取出某一天的逐時資料。
    需求欄位：col_map['datetime'], col_map['value']
    回傳：DataFrame(columns=['time','value'])，time 為 '00:00' ~ '23:00' 字串
    """
    _ensure_exists(xlsx)
    df = pd.read_excel(xlsx, sheet_name=sheet)

    dt_col = col_map["datetime"]
    v_col  = col_map["value"]

    if dt_col not in df.columns or v_col not in df.columns:
        raise ValueError(
            f"資料格式錯誤：欄位找不到（需要 {dt_col}/{v_col}）。目前欄位：{list(df.columns)}"
        )

    # 解析 datetime
    df[dt_col] = pd.to_datetime(df[dt_col], errors="coerce")
    if df[dt_col].isna().all():
        raise ValueError("日期欄解析失敗，請確認格式")

    # 過濾到指定那一天
    target_date = pd.to_datetime(date_str).date()
    df_day = df[df[dt_col].dt.date == target_date].copy()

    if df_day.empty:
        # 回傳空 DataFrame，前端會顯示「查無資料」
        return pd.DataFrame(columns=["time", "value"])

    df_day["time"] = df_day[dt_col].dt.strftime("%H:%M")
    df_day["value"] = pd.to_numeric(df_day[v_col], errors="coerce").fillna(0.0)

    # 只留必要欄位並依時間排序
    df_day = df_day[["time", "value"]].sort_values("time", kind="stable").reset_index(drop=True)
    return df_day

INDEX_HTML = """
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>中西區 里別負載 + PV 查詢</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1"></script>
</head>
<body class="bg-light">
<div class="container py-4">
  <h3 class="mb-4">中西區：選里＋日期 → 顯示「一天負載」與「同日 PV 發電」</h3>

  <form class="row gy-2 gx-3 align-items-end mb-4" method="get" action="{{ url_for('view') }}">
    <div class="col-auto">
      <label class="form-label">里名</label>
      <select class="form-select" name="village" required>
        <option value="" disabled selected>請選擇</option>
        {% for v in villages %}
          <option value="{{v}}">{{v}}</option>
        {% endfor %}
      </select>
    </div>
    <div class="col-auto">
      <label class="form-label">日期</label>
      <input type="date" class="form-control" name="date" required>
    </div>
    <div class="col-auto">
      <button class="btn btn-primary">查詢</button>
    </div>
  </form>

  <p class="text-muted">資料來源：<code>{{load_path}}</code>（里別負載），<code>{{pv_path}}</code>（中西區 PV）</p>
</div>
</body>
</html>
"""

VIEW_HTML = """
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>{{ village }} - {{ date }} | 負載 & PV</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1"></script>
</head>
<body class="bg-light">
<div class="container py-4">
  <div class="d-flex align-items-center mb-3">
    <a class="btn btn-outline-secondary me-3" href="{{ url_for('index') }}">← 返回</a>
    <h3 class="m-0">{{ village }} / {{ date }}</h3>
  </div>

  {% if error %}
    <div class="alert alert-danger">{{ error }}</div>
  {% endif %}

  <div class="row g-4">
    <!-- 里別負載 -->
    <div class="col-12 col-lg-6">
      <div class="card shadow-sm">
        <div class="card-header fw-bold">里別 24 小時負載（MW）</div>
        <div class="card-body">
          {% if load_labels %}
            <canvas id="loadChart"></canvas>
          {% else %}
            <div class="text-muted">查無資料</div>
          {% endif %}
        </div>
        <div class="card-body pt-0">
          {% if load_rows %}
            <div class="table-responsive">
              <table class="table table-sm table-striped align-middle">
                <thead><tr><th>時間</th><th class="text-end">負載 (MW)</th></tr></thead>
                <tbody>
                {% for r in load_rows %}
                  <tr><td>{{ r.time }}</td><td class="text-end">{{ "%.6f"|format(r.value) }}</td></tr>
                {% endfor %}
                </tbody>
              </table>
            </div>
          {% endif %}
        </div>
      </div>
    </div>

    <!-- 中西區 PV -->
    <div class="col-12 col-lg-6">
      <div class="card shadow-sm">
        <div class="card-header fw-bold">中西區同日 PV 發電（MW）</div>
        <div class="card-body">
          {% if pv_labels %}
            <canvas id="pvChart"></canvas>
          {% else %}
            <div class="text-muted">查無資料</div>
          {% endif %}
        </div>
        <div class="card-body pt-0">
          {% if pv_rows %}
            <div class="table-responsive">
              <table class="table table-sm table-striped align-middle">
                <thead><tr><th>時間</th><th class="text-end">PV (MW)</th></tr></thead>
                <tbody>
                {% for r in pv_rows %}
                  <tr><td>{{ r.time }}</td><td class="text-end">{{ "%.6f"|format(r.value) }}</td></tr>
                {% endfor %}
                </tbody>
              </table>
            </div>
          {% endif %}
        </div>
      </div>
    </div>
  </div>
</div>

<script>
  {% if load_labels %}
  new Chart(document.getElementById('loadChart'), {
    type: 'line',
    data: {
      labels: {{ load_labels|tojson }},
      datasets: [{label: 'Load (MW)', data: {{ load_values|tojson }}, fill: false, tension: 0.2}]
    },
    options: {responsive: true, scales: {y: {beginAtZero: true}}}
  });
  {% endif %}

  {% if pv_labels %}
  new Chart(document.getElementById('pvChart'), {
    type: 'line',
    data: {
      labels: {{ pv_labels|tojson }},
      datasets: [{label: 'PV (MW)', data: {{ pv_values|tojson }}, fill: false, tension: 0.2}]
    },
    options: {responsive: true, scales: {y: {beginAtZero: true}}}
  });
  {% endif %}
</script>
</body>
</html>
"""

@app.route("/")
def index():
    try:
        villages = list_villages_from_load()
    except Exception as e:
        villages = []
        print("讀取 villages 失敗:", e)

    return render_template_string(
        INDEX_HTML,
        villages=villages,
        load_path=str(LOAD_XLSX),
        pv_path=str(PV_XLSX),
    )

@app.route("/view")
def view():
    village = request.args.get("village", "").strip()
    date_str = request.args.get("date", "").strip()
    if not village or not date_str:
        return redirect(url_for("index"))

    error_msg = None

    # 里別負載
    try:
        df_load = day_series_from_sheet(LOAD_XLSX, village, date_str, LOAD_COLS)
        load_labels = df_load["time"].tolist()
        load_values = df_load["value"].round(6).tolist()
        load_rows   = df_load.to_dict(orient="records")
    except Exception as e:
        error_msg = f"讀取里別負載失敗：{e}"
        load_labels = load_values = []
        load_rows   = []

    # 同日 PV
    try:
        df_pv = day_series_from_sheet(PV_XLSX, PV_SHEET, date_str, PV_COLS)
        pv_labels = df_pv["time"].tolist()
        pv_values = df_pv["value"].round(6).tolist()
        pv_rows   = df_pv.to_dict(orient="records")
    except Exception as e:
        # 不讓整頁爆掉；只在上面顯示錯誤
        error_msg = (error_msg + "；" if error_msg else "") + f"讀取 PV 失敗：{e}"
        pv_labels = pv_values = []
        pv_rows   = []

    return render_template_string(
        VIEW_HTML,
        village=village, date=date_str,
        load_labels=load_labels, load_values=load_values, load_rows=load_rows,
        pv_labels=pv_labels, pv_values=pv_values, pv_rows=pv_rows,
        error=error_msg
    )

# for Render
app = app

if __name__ == "__main__":
    # 本機執行
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
