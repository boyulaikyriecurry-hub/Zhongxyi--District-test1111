from __future__ import annotations
from pathlib import Path
import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for

# ========= 可調整區 =========
DATA_DIR = Path(__file__).parent / "data"

DATA_XLSX = DATA_DIR / "放到網頁的data.xlsx"   # 只吃這一個檔案

# 里別 sheet 內欄位名稱
DATA_COLS = {
    "datetime": "datetime",
    "load": "load_kWh",
    "critical_load": "critical_load_kWh",
    "pv": "pv_kWh",
    "soc": "SOC_kWh",
    "bess_capacity": "bess_capacity_kwh",
    "pv_capacity": "pv_capacity_kwh",
}
# ========= 可調整區 =========

app = Flask(__name__)

def _ensure_exists(p: Path):
    if not p.exists():
        raise FileNotFoundError(f"找不到資料檔：{p}")

def list_villages() -> list[str]:
    """列出資料檔中可用的里名（工作表名稱）。"""
    _ensure_exists(DATA_XLSX)
    xls = pd.ExcelFile(DATA_XLSX)

    # 過濾掉不是里名的 sheet
    exclude_sheets = {
        "Summary",
        "ALL_LI_TOTAL",
        "Installed_Capacity_By_LI",
        "Installed_Capacity_Summary",
    }
    villages = [s for s in xls.sheet_names if s not in exclude_sheets]
    return villages

def read_day_data(village: str, date_str: str) -> pd.DataFrame:
    """
    讀取某個里的某一天資料。
    回傳欄位包含：
    time, load_kWh, critical_load_kWh, pv_kWh, SOC_kWh
    """
    _ensure_exists(DATA_XLSX)
    df = pd.read_excel(DATA_XLSX, sheet_name=village)

    dt_col = DATA_COLS["datetime"]
    load_col = DATA_COLS["load"]
    critical_col = DATA_COLS["critical_load"]
    pv_col = DATA_COLS["pv"]
    soc_col = DATA_COLS["soc"]

    required_cols = [dt_col, load_col, critical_col, pv_col, soc_col]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        raise ValueError(f"資料格式錯誤，缺少欄位：{missing_cols}。目前欄位：{list(df.columns)}")

    df[dt_col] = pd.to_datetime(df[dt_col], errors="coerce")
    if df[dt_col].isna().all():
        raise ValueError("datetime 欄位解析失敗，請確認格式")

    target_date = pd.to_datetime(date_str).date()
    df_day = df[df[dt_col].dt.date == target_date].copy()

    if df_day.empty:
        return pd.DataFrame(columns=[
            "time",
            load_col,
            critical_col,
            pv_col,
            soc_col,
        ])

    df_day["time"] = df_day[dt_col].dt.strftime("%H:%M")

    numeric_cols = [load_col, critical_col, pv_col, soc_col]
    for col in numeric_cols:
        df_day[col] = pd.to_numeric(df_day[col], errors="coerce").fillna(0.0)

    df_day = df_day[["time", load_col, critical_col, pv_col, soc_col]].sort_values(
        "time", kind="stable"
    ).reset_index(drop=True)

    return df_day

def read_capacity_info(village: str) -> tuple[float, float]:
    """
    讀取某個里的 bess_capacity_kwh 與 pv_capacity_kwh。
    預設取該 sheet 第一列。
    """
    _ensure_exists(DATA_XLSX)
    df = pd.read_excel(DATA_XLSX, sheet_name=village)

    bess_col = DATA_COLS["bess_capacity"]
    pv_cap_col = DATA_COLS["pv_capacity"]

    bess_capacity = 0.0
    pv_capacity = 0.0

    if bess_col in df.columns and not df.empty:
        bess_capacity = pd.to_numeric(df[bess_col], errors="coerce").fillna(0.0).iloc[0]

    if pv_cap_col in df.columns and not df.empty:
        pv_capacity = pd.to_numeric(df[pv_cap_col], errors="coerce").fillna(0.0).iloc[0]

    return float(bess_capacity), float(pv_capacity)

INDEX_HTML = """
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>中西區里別關鍵負載查詢</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light" style="
       background-image: url('{{ url_for('static', filename='img_160854413174.jpg') }}');
       background-size: cover;
       background-position: center;
     ">
<div class="container py-4">
  <h3 class="mb-4">中西區：選里＋日期</h3>

  <form class="row gy-2 gx-3 align-items-end mb-4" method="get" action="{{ url_for('view') }}">
    <div class="col-auto">
      <label class="form-label">里名</label>
      <select class="form-select" name="village" required>
        <option value="" disabled selected>請選擇</option>
        {% for v in villages %}
          <option value="{{ v }}">{{ v }}</option>
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

  <p class="text-muted">資料來源：<code>{{ data_path }}</code></p>
</div>

<div class="container py-4">
  <img src="{{ url_for('static', filename='中西區調整後地圖-簡.jpg') }}"
       alt="map"
       height="1000"
       class="img-fluid mb-4">
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
  <title>{{ village }} - {{ date }} | 里別資料</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1"></script>
</head>
<body class="bg-light">
<div class="container-fluid py-4">

  <div class="d-flex align-items-center mb-3">
    <a class="btn btn-outline-secondary me-3" href="{{ url_for('index') }}">← 返回</a>
    <h3 class="m-0">{{ village }} / {{ date }}</h3>
  </div>

  <div class="mb-3">
    <a class="btn btn-outline-primary me-2"
       href="{{ url_for('view', village=village, date=date, mode='chart') }}">
       圖表模式
    </a>
    <a class="btn btn-outline-secondary"
       href="{{ url_for('view', village=village, date=date, mode='table') }}">
       表格模式
    </a>
  </div>

  {% if error %}
    <div class="alert alert-danger">{{ error }}</div>
  {% endif %}

  {% if mode == 'chart' %}
  <div class="row g-3 mb-4">

    <div class="col-12 col-lg-6">
      <div class="card shadow-sm h-100">
        <div class="card-header fw-bold text-center">關鍵負載</div>
        <div class="card-body">
          <canvas id="criticalLoadChart"></canvas>
        </div>
      </div>
    </div>

    <div class="col-12 col-lg-6">
      <div class="card shadow-sm h-100">
        <div class="card-header fw-bold text-center">PV 發電量</div>
        <div class="card-body">
          <canvas id="pvChart"></canvas>
        </div>
      </div>
    </div>
    <div class="col-12 col-lg-6">
      <div class="card shadow-sm h-100">
        <div class="card-header fw-bold text-center">負載</div>
        <div class="card-body">
          <canvas id="loadChart"></canvas>
        </div>
      </div>
    </div>
    <div class="col-12 col-lg-6">
      <div class="card shadow-sm h-100">
        <div class="card-header fw-bold text-center">電池容量</div>
        <div class="card-body">
          <canvas id="socChart"></canvas>
        </div>
      </div>
    </div>
  </div>
  {% endif %}

  {% if mode == 'table' %}
  <div class="row g-3">
    <div class="col-12">
      <div class="card shadow-sm">
        <div class="card-header fw-bold text-center">逐時資料表</div>
        <div class="card-body table-responsive">
          <table class="table table-sm table-striped align-middle">
            <thead>
              <tr>
                <th>時間</th>
                <th class="text-end">負載 (kWh)</th>
                <th class="text-end">關鍵負載 (kWh)</th>
                <th class="text-end">PV (kWh)</th>
                <th class="text-end">SOC (kWh)</th>
              </tr>
            </thead>
            <tbody>
            {% for r in detail_rows %}
              <tr>
                <td>{{ r.time }}</td>
                <td class="text-end">{{ "%.6f"|format(r.load_kWh) }}</td>
                <td class="text-end">{{ "%.6f"|format(r.critical_load_kWh) }}</td>
                <td class="text-end">{{ "%.6f"|format(r.pv_kWh) }}</td>
                <td class="text-end">{{ "%.6f"|format(r.SOC_kWh) }}</td>
              </tr>
            {% endfor %}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
  {% endif %}

</div>

<script>
  {% if mode == 'chart' and time_labels %}

  new Chart(document.getElementById('criticalLoadChart'), {
    type: 'line',
    data: {
      labels: {{ time_labels|tojson }},
      datasets: [
        {
          label: '關鍵負載 (kWh)',
          borderColor: 'rgba(231, 76, 60, 1)',   // 紅色
          data: {{ critical_load_values|tojson }},
          fill: false,
          tension: 0.2
        }
      ]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('pvChart'), {
    type: 'line',
    data: {
      labels: {{ time_labels|tojson }},
      datasets: [
        {
          label: 'PV (kWh)',
          borderColor: 'rgba(46, 204, 113, 1)',  // 綠色
          data: {{ pv_values|tojson }},
          fill: false,
          tension: 0.2
        }
      ]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
  new Chart(document.getElementById('loadChart'), {
  type: 'line',
  data: {
    labels: {{ time_labels|tojson }},
    datasets: [
      {
        label: '負載 (kWh)',
        borderColor: 'rgba(52, 152, 219, 1)',  // 藍色
        data: {{ load_values|tojson }},
        fill: false,
        tension: 0.2
      }
    ]
  },
  options: {
    responsive: true,
    scales: {
      y: { beginAtZero: true }
    }
  }
});
new Chart(document.getElementById('socChart'), {
  type: 'line',
  data: {
    labels: {{ time_labels|tojson }},
    datasets: [
      {
        label: 'SOC (kWh)',
        borderColor: 'rgba(243, 156, 18, 1)',
        data: {{ soc_values|tojson }},
        fill: false,
        tension: 0.2
      }
    ]
  },
  options: {
    responsive: true,
    scales: {
      y: { beginAtZero: true }
    }
  }
});
  {% endif %}
</script>
</body>
</html>
"""

@app.route("/")
def index():
    try:
        villages = list_villages()   # 改成新的函式
    except Exception as e:
        villages = []
        print("讀取 villages 失敗:", e)

    return render_template_string(
        INDEX_HTML,
        villages=villages,
        data_path=str(DATA_XLSX),
    )


@app.route("/view")
def view():
    village = request.args.get("village", "").strip()
    date_str = request.args.get("date", "").strip()
    mode = request.args.get("mode", "chart").strip()

    if not village or not date_str:
        return redirect(url_for("index"))

    error_msg = None

    time_labels = []
    load_values = []
    critical_load_values = []
    pv_values = []
    soc_values = []
    detail_rows = []

    try:
        df_day = read_day_data(village, date_str)

        if not df_day.empty:
            time_labels = df_day["time"].tolist()
            load_values = df_day[DATA_COLS["load"]].round(6).tolist()
            critical_load_values = df_day[DATA_COLS["critical_load"]].round(6).tolist()
            pv_values = df_day[DATA_COLS["pv"]].round(6).tolist()
            soc_values = df_day[DATA_COLS["soc"]].round(6).tolist()

            detail_rows = df_day.to_dict(orient="records")
        else:
            error_msg = f"{village} 在 {date_str} 查無資料"

    except Exception as e:
        error_msg = f"讀取 {village} 資料失敗：{e}"

    return render_template_string(
        VIEW_HTML,
        village=village,
        date=date_str,
        mode=mode,
        error=error_msg,
        time_labels=time_labels,
        load_values=load_values,
        critical_load_values=critical_load_values,
        pv_values=pv_values,
        soc_values=soc_values,
        detail_rows=detail_rows,
    )

# for Render
app = app

if __name__ == "__main__":
    # 本機執行
    app.run(debug=True)
