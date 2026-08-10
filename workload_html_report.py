import json
import os

import numpy as np
import pandas as pd

from workload_excel_report import WorkloadExcelReporter


class WorkloadHtmlVisualizer:

    METRICS = [
        "Planned Start Volume",
        "Received Start Volume",
        "WIP Received Start Volume",
        "Remaining Planned Start Volume",
        "Actual Start Volume",
        "Actual Completion Volume",
        "Forecast Completion Volume",
        "Completion Volume",
        "Init Volume",
        "Demand FTE",
        "Forecast WBH Letter Volume",
        "Forecast WBH Call Volume",
    ]

    DEFAULT_METRICS = [
        "Planned Start Volume",
        "Received Start Volume",
        "Remaining Planned Start Volume",
        "Actual Completion Volume",
        "Forecast Completion Volume",
        "Demand FTE",
    ]

    FREQUENCY_LABELS = {
        "D": "Daily",
        "W": "Weekly",
        "M": "Monthly",
    }

    DETAIL_TABLE_ORDER = [
        "02_Workload_Long",
        "12_Start_Reconciliation",
        "13_WIP_Received_Start",
        "20_Completion",
        "21_Completion_Distribution",
        "30_WBH_Letter",
        "31_WBH_Call",
        "40_Calc_Forecast_Completion",
        "41_Calc_WBH_Letter",
        "42_Calc_WBH_Call",
        "43_Calc_Demand_FTE",
        "00_Control",
        "03_Metric_Definitions",
        "10_Input_BoW",
        "11_Planned_Start",
    ]

    def __init__(self, processors_by_frequency, output_path=None, max_detail_rows=5000):
        if not processors_by_frequency:
            raise ValueError("processors_by_frequency is required.")

        self.processors_by_frequency = processors_by_frequency
        self.output_path = output_path or os.path.join("data", "Workload_Forecast_Visualization.html")
        self.max_detail_rows = max_detail_rows
        self.visualization_data = {}

    def build_visualization_data(self):
        frequencies = {}
        metric_definitions = None

        for frequency, processor in self.processors_by_frequency.items():
            reporter = WorkloadExcelReporter(processor)
            output_tables = reporter.build_output_excel_tables(write_excel=False)
            output_tables["40_Calc_Forecast_Completion"] = (
                reporter.build_forecast_completion_calculation_table()
            )
            output_tables["41_Calc_WBH_Letter"] = reporter.build_wbh_action_calculation_table("WBH Letter")
            output_tables["42_Calc_WBH_Call"] = reporter.build_wbh_action_calculation_table("WBH Call")
            output_tables["43_Calc_Demand_FTE"] = reporter.build_demand_fte_calculation_table()

            if metric_definitions is None:
                metric_definitions = output_tables.get("03_Metric_Definitions", pd.DataFrame())

            frequencies[frequency] = self.build_frequency_payload(
                frequency,
                processor,
                output_tables
            )

        self.visualization_data = {
            "title": "CDD Workload Forecast",
            "metrics": self.METRICS,
            "defaultMetrics": self.DEFAULT_METRICS,
            "metricDefinitions": self.table_to_records(metric_definitions),
            "frequencyLabels": self.FREQUENCY_LABELS,
            "frequencies": frequencies,
        }
        return self.visualization_data

    def build_frequency_payload(self, frequency, processor, output_tables):
        workload_long = output_tables.get("02_Workload_Long", pd.DataFrame())
        workload_long = workload_long[workload_long["Metric"].isin(self.METRICS)].copy()

        if workload_long.empty:
            periods = []
            case_types = []
        else:
            periods_df = workload_long[["Period", "Period Start", "Period End"]].drop_duplicates()
            periods_df = periods_df.sort_values(["Period Start", "Period"])
            periods = self.table_to_records(periods_df)
            case_types = sorted(workload_long["Case Type"].dropna().astype(str).unique())

        details = {}
        detail_counts = {}
        for table_name in self.DETAIL_TABLE_ORDER:
            if table_name in output_tables:
                detail_df = output_tables[table_name]
                total_rows = 0 if detail_df is None else len(detail_df)
                detail_records = self.table_to_records(detail_df, max_rows=self.max_detail_rows)
                details[table_name] = detail_records
                detail_counts[table_name] = {
                    "includedRows": len(detail_records),
                    "totalRows": total_rows,
                }

        control_table = output_tables.get("00_Control", pd.DataFrame())
        control = {}
        if not control_table.empty and {"Item", "Value"}.issubset(control_table.columns):
            for _, row in control_table.iterrows():
                control[str(row["Item"])] = self.clean_value(row["Value"])

        return {
            "frequency": frequency,
            "label": self.FREQUENCY_LABELS.get(frequency, frequency),
            "periods": periods,
            "caseTypes": case_types,
            "control": control,
            "workload": self.table_to_records(workload_long),
            "details": details,
            "detailCounts": detail_counts,
            "detailLabels": self.build_detail_labels(details, detail_counts),
        }

    def build_detail_labels(self, details, detail_counts):
        labels = {}
        for table_name, rows in details.items():
            clean_name = table_name.split("_", 1)[-1].replace("_", " ")
            counts = detail_counts.get(table_name, {})
            total_rows = counts.get("totalRows", len(rows))
            included_rows = counts.get("includedRows", len(rows))
            if included_rows < total_rows:
                row_label = f"{included_rows:,}/{total_rows:,} rows"
            else:
                row_label = f"{total_rows:,} rows"
            labels[table_name] = f"{clean_name} ({row_label})"
        return labels

    def table_to_records(self, output_df, max_rows=None):
        if output_df is None or output_df.empty:
            return []

        output_df = output_df.copy()
        if max_rows is not None:
            output_df = output_df.head(max_rows)
        return [
            {
                str(column): self.clean_value(value)
                for column, value in row.items()
            }
            for row in output_df.to_dict(orient="records")
        ]

    def clean_value(self, value):
        if value is None:
            return None
        if isinstance(value, pd.Period):
            return str(value)
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return None
            return value.strftime("%Y-%m-%d")
        if isinstance(value, np.integer):
            return int(value)
        if isinstance(value, np.floating):
            if not np.isfinite(value):
                return None
            return float(value)
        if isinstance(value, float):
            if not np.isfinite(value):
                return None
            return value
        if pd.isna(value):
            return None
        return value

    def build_html(self):
        if not self.visualization_data:
            self.build_visualization_data()

        payload = json.dumps(self.visualization_data, ensure_ascii=False).replace("</", "<\\/")
        return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CDD Workload Forecast</title>
  <style>
    :root {{
      --bg: #f7f8fa;
      --surface: #ffffff;
      --surface-2: #eef3f4;
      --line: #d9e0e3;
      --ink: #1d2528;
      --muted: #627174;
      --teal: #0b6f6a;
      --blue: #315f9f;
      --gold: #b36a00;
      --rose: #b23b53;
      --green: #3f7d3f;
      --violet: #7456a3;
      --radius: 8px;
    }}

    * {{
      box-sizing: border-box;
    }}

    body {{
      margin: 0;
      color: var(--ink);
      background: var(--bg);
      font-family: "Segoe UI", Arial, sans-serif;
      font-size: 14px;
      letter-spacing: 0;
    }}

    header {{
      border-bottom: 1px solid var(--line);
      background: var(--surface);
    }}

    .header-inner {{
      max-width: 1480px;
      margin: 0 auto;
      padding: 18px 24px 14px;
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 20px;
    }}

    h1 {{
      margin: 0;
      font-size: 24px;
      font-weight: 700;
      line-height: 1.2;
    }}

    .meta {{
      display: flex;
      flex-wrap: wrap;
      justify-content: flex-end;
      gap: 8px;
      color: var(--muted);
      font-size: 12px;
    }}

    .meta span {{
      padding: 5px 8px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface-2);
      white-space: nowrap;
    }}

    main {{
      max-width: 1480px;
      margin: 0 auto;
      padding: 18px 24px 30px;
    }}

    section {{
      margin-bottom: 18px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface);
    }}

    .section-head {{
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }}

    h2 {{
      margin: 0;
      font-size: 15px;
      font-weight: 700;
    }}

    .filters {{
      padding: 14px 16px;
      display: grid;
      grid-template-columns: minmax(180px, 0.8fr) minmax(260px, 2fr) minmax(180px, 1fr);
      gap: 16px;
      align-items: start;
    }}

    .field-label {{
      display: block;
      margin-bottom: 8px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
    }}

    .segmented {{
      display: inline-flex;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      overflow: hidden;
      background: var(--surface);
    }}

    .segmented button {{
      border: 0;
      border-right: 1px solid var(--line);
      padding: 8px 12px;
      background: transparent;
      color: var(--ink);
      cursor: pointer;
      font: inherit;
    }}

    .segmented button:last-child {{
      border-right: 0;
    }}

    .segmented button.active {{
      background: var(--teal);
      color: #fff;
    }}

    .metric-list,
    .case-list {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }}

    label.choice {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 7px 9px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: #fff;
      cursor: pointer;
      white-space: nowrap;
    }}

    label.choice input {{
      margin: 0;
    }}

    .toolbar {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }}

    .plain-button,
    select,
    input[type="search"] {{
      border: 1px solid var(--line);
      border-radius: var(--radius);
      padding: 8px 10px;
      background: #fff;
      color: var(--ink);
      font: inherit;
    }}

    .plain-button {{
      cursor: pointer;
    }}

    .plain-button:disabled {{
      cursor: not-allowed;
      opacity: 0.48;
    }}

    .plain-button.primary {{
      background: var(--teal);
      border-color: var(--teal);
      color: #fff;
    }}

    .content {{
      padding: 14px 16px 16px;
    }}

    .summary-grid {{
      display: grid;
      grid-template-columns: repeat(4, minmax(160px, 1fr));
      gap: 10px;
    }}

    .tile {{
      border: 1px solid var(--line);
      border-radius: var(--radius);
      padding: 11px 12px;
      background: #fff;
      min-height: 88px;
    }}

    .tile-title {{
      color: var(--muted);
      font-size: 12px;
      line-height: 1.25;
      min-height: 30px;
    }}

    .tile-value {{
      margin-top: 8px;
      font-size: 22px;
      font-weight: 700;
    }}

    .tile-sub {{
      margin-top: 4px;
      color: var(--muted);
      font-size: 12px;
    }}

    .chart-wrap {{
      position: relative;
      min-height: 410px;
    }}

    .time-controls {{
      margin-bottom: 10px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      padding: 10px 12px;
      background: #fbfcfc;
    }}

    .time-toolbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }}

    .time-range-row {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}

    .time-range-row input[type="range"] {{
      flex: 1;
      min-width: 160px;
      accent-color: var(--teal);
    }}

    .time-window-text {{
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
    }}

    svg {{
      display: block;
      width: 100%;
      height: 390px;
    }}

    .legend {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      color: var(--muted);
      font-size: 12px;
      padding-top: 8px;
    }}

    .legend-item {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }}

    .swatch {{
      width: 11px;
      height: 11px;
      border-radius: 2px;
      display: inline-block;
    }}

    .tooltip {{
      position: fixed;
      pointer-events: none;
      z-index: 10;
      min-width: 180px;
      max-width: 280px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      padding: 8px 10px;
      background: #fff;
      box-shadow: 0 8px 20px rgba(21, 32, 36, 0.16);
      color: var(--ink);
      font-size: 12px;
      display: none;
    }}

    .table-wrap {{
      overflow: auto;
      max-height: 520px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
      background: #fff;
      font-size: 12px;
    }}

    th,
    td {{
      padding: 7px 9px;
      border-bottom: 1px solid #e8edef;
      border-right: 1px solid #edf1f2;
      text-align: left;
      white-space: nowrap;
    }}

    th {{
      position: sticky;
      top: 0;
      z-index: 1;
      background: #21353a;
      color: #fff;
      font-weight: 700;
    }}

    tr:nth-child(even) td {{
      background: #fbfcfc;
    }}

    .matrix-grid {{
      width: max-content;
      min-width: 100%;
      border-collapse: separate;
      border-spacing: 0;
    }}

    .matrix-grid .metric-col {{
      left: 0;
      min-width: 220px;
      max-width: 220px;
    }}

    .matrix-grid .case-col {{
      left: 220px;
      min-width: 120px;
      max-width: 120px;
      box-shadow: 5px 0 7px rgba(29, 37, 40, 0.08);
    }}

    .matrix-grid .period-col {{
      min-width: 96px;
    }}

    .matrix-grid .sticky-col {{
      position: sticky;
      z-index: 3;
      background: #fff;
    }}

    .matrix-grid th.sticky-col {{
      z-index: 5;
      background: #21353a;
    }}

    .matrix-grid tr:nth-child(even) td.sticky-col {{
      background: #fbfcfc;
    }}

    .muted {{
      color: var(--muted);
    }}

    .empty {{
      padding: 34px;
      color: var(--muted);
      text-align: center;
    }}

    @media (max-width: 900px) {{
      .header-inner {{
        align-items: flex-start;
        flex-direction: column;
      }}

      .meta {{
        justify-content: flex-start;
      }}

      .filters {{
        grid-template-columns: 1fr;
      }}

      .summary-grid {{
        grid-template-columns: repeat(2, minmax(150px, 1fr));
      }}
    }}

    @media (max-width: 560px) {{
      main,
      .header-inner {{
        padding-left: 14px;
        padding-right: 14px;
      }}

      .summary-grid {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="header-inner">
      <div>
        <h1>CDD Workload Forecast</h1>
        <div class="muted" id="subtitle"></div>
      </div>
      <div class="meta" id="meta"></div>
    </div>
  </header>

  <main>
    <section>
      <div class="filters">
        <div>
          <span class="field-label">Frequency</span>
          <div class="segmented" id="frequencyButtons"></div>
          <div style="margin-top:12px">
            <span class="field-label">Chart Scale</span>
            <div class="segmented" id="scaleButtons">
              <button type="button" data-scale="absolute" class="active">Absolute</button>
              <button type="button" data-scale="indexed">Indexed</button>
            </div>
          </div>
        </div>
        <div>
          <span class="field-label">Metrics</span>
          <div class="toolbar" style="margin-bottom:8px">
            <button class="plain-button primary" id="selectCore" type="button">Core Set</button>
            <button class="plain-button" id="selectAllMetrics" type="button">Select All</button>
            <button class="plain-button" id="clearMetrics" type="button">Clear</button>
          </div>
          <div class="metric-list" id="metricList"></div>
        </div>
        <div>
          <span class="field-label">Case Type</span>
          <div class="toolbar" style="margin-bottom:8px">
            <button class="plain-button primary" id="selectAllCases" type="button">All Case Types</button>
          </div>
          <div class="case-list" id="caseList"></div>
        </div>
      </div>
    </section>

    <section>
      <div class="section-head">
        <h2>Selected Period Summary</h2>
        <span class="muted" id="summaryNote"></span>
      </div>
      <div class="content">
        <div class="summary-grid" id="summaryGrid"></div>
      </div>
    </section>

    <section>
      <div class="section-head">
        <h2>Trend</h2>
        <span class="muted" id="chartNote"></span>
      </div>
      <div class="content">
        <div class="time-controls">
          <div class="time-toolbar">
            <div class="toolbar">
              <button class="plain-button" id="timeZoomIn" type="button">Zoom In</button>
              <button class="plain-button" id="timeZoomOut" type="button">Zoom Out</button>
              <button class="plain-button" id="timeZoomReset" type="button">Full Range</button>
            </div>
            <span class="time-window-text" id="timeWindowText"></span>
          </div>
          <div class="time-range-row">
            <span class="muted">Start</span>
            <input id="timeWindowSlider" type="range" min="0" max="0" value="0" step="1">
            <span class="muted">End</span>
          </div>
        </div>
        <div class="chart-wrap">
          <svg id="trendChart" viewBox="0 0 1100 390" role="img"></svg>
          <div class="tooltip" id="tooltip"></div>
        </div>
        <div class="legend" id="legend"></div>
      </div>
    </section>

    <section>
      <div class="section-head">
        <h2>Metric Matrix</h2>
        <span class="muted" id="matrixNote"></span>
      </div>
      <div class="content">
        <div class="table-wrap" id="matrixTable"></div>
      </div>
    </section>

    <section>
      <div class="section-head">
        <h2>Calculation Details</h2>
        <div class="toolbar">
          <select id="detailTableSelect"></select>
          <input type="search" id="detailSearch" placeholder="Search details">
        </div>
      </div>
      <div class="content">
        <div class="muted" id="detailNote" style="margin-bottom:8px"></div>
        <div class="table-wrap" id="detailTable"></div>
      </div>
    </section>
  </main>

  <div class="tooltip" id="pointTooltip"></div>

  <script>
    const DATA = {payload};
    const COLORS = [
      "#0b6f6a", "#315f9f", "#b36a00", "#b23b53", "#3f7d3f", "#7456a3",
      "#6f4a22", "#657b83", "#a34f82", "#257a9b", "#8a6f16", "#516f3b"
    ];

    const state = {{
      frequency: Object.keys(DATA.frequencies).includes("M") ? "M" : Object.keys(DATA.frequencies)[0],
      metrics: new Set(DATA.defaultMetrics),
      caseTypes: new Set(),
      scale: "absolute",
      timeWindows: {{}},
      detailTable: "02_Workload_Long",
      detailSearch: ""
    }};

    function esc(value) {{
      return String(value ?? "").replace(/[&<>"']/g, ch => ({{
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      }}[ch]));
    }}

    function currentData() {{
      return DATA.frequencies[state.frequency];
    }}

    function metricColor(metric) {{
      const index = DATA.metrics.indexOf(metric);
      return COLORS[(index >= 0 ? index : 0) % COLORS.length];
    }}

    function formatNumber(value, key = "") {{
      if (value === null || value === undefined || value === "") return "";
      if (typeof value !== "number") return value;
      if (key.includes("Probability")) return `${{(value * 100).toFixed(1)}}%`;
      if (key.includes("FTE")) return value.toLocaleString(undefined, {{ maximumFractionDigits: 2, minimumFractionDigits: 2 }});
      if (key.includes("UPT") || key.includes("Hour")) return value.toLocaleString(undefined, {{ maximumFractionDigits: 1 }});
      return value.toLocaleString(undefined, {{ maximumFractionDigits: 1 }});
    }}

    function getSelectedRows() {{
      const data = currentData();
      return data.workload.filter(row =>
        state.metrics.has(row.Metric) && state.caseTypes.has(row["Case Type"])
      );
    }}

    function aggregateRows(rows) {{
      const byMetricPeriod = new Map();
      const periodSet = new Set();
      for (const row of rows) {{
        const key = `${{row.Metric}}||${{row.Period}}`;
        byMetricPeriod.set(key, (byMetricPeriod.get(key) || 0) + (Number(row.Volume) || 0));
        periodSet.add(row.Period);
      }}
      const allPeriods = currentData().periods.map(period => period.Period).filter(period => periodSet.has(period));
      return {{ byMetricPeriod, periods: allPeriods }};
    }}

    function clamp(value, min, max) {{
      return Math.max(min, Math.min(max, value));
    }}

    function getTimeWindow(totalPeriods) {{
      const frequency = state.frequency;
      const total = Math.max(0, totalPeriods);
      if (!state.timeWindows[frequency]) {{
        state.timeWindows[frequency] = {{ start: 0, size: total }};
      }}

      const windowState = state.timeWindows[frequency];
      windowState.size = clamp(
        Math.round(windowState.size || total || 1),
        total > 1 ? 2 : total,
        Math.max(total, 1)
      );
      if (total <= 1) {{
        windowState.start = 0;
        windowState.size = total;
        return windowState;
      }}

      windowState.start = clamp(
        Math.round(windowState.start || 0),
        0,
        Math.max(0, total - windowState.size)
      );
      return windowState;
    }}

    function setTimeWindowSize(totalPeriods, newSize) {{
      const windowState = getTimeWindow(totalPeriods);
      const total = Math.max(0, totalPeriods);
      if (total <= 1) return windowState;

      const oldCenter = windowState.start + (windowState.size - 1) / 2;
      windowState.size = clamp(Math.round(newSize), 2, total);
      windowState.start = clamp(
        Math.round(oldCenter - (windowState.size - 1) / 2),
        0,
        Math.max(0, total - windowState.size)
      );
      return windowState;
    }}

    function getVisiblePeriods(periods) {{
      const windowState = getTimeWindow(periods.length);
      return periods.slice(windowState.start, windowState.start + windowState.size);
    }}

    function renderTimeControls(periods) {{
      const total = periods.length;
      const windowState = getTimeWindow(total);
      const maxStart = Math.max(0, total - windowState.size);
      const slider = document.getElementById("timeWindowSlider");
      const label = document.getElementById("timeWindowText");
      const zoomIn = document.getElementById("timeZoomIn");
      const zoomOut = document.getElementById("timeZoomOut");
      const reset = document.getElementById("timeZoomReset");

      slider.min = 0;
      slider.max = maxStart;
      slider.value = windowState.start;
      slider.disabled = maxStart === 0;
      zoomIn.disabled = total <= 2 || windowState.size <= 2;
      zoomOut.disabled = total <= 1 || windowState.size >= total;
      reset.disabled = total <= 1 || (windowState.start === 0 && windowState.size >= total);

      const startPeriod = periods[windowState.start] || "";
      const endPeriod = periods[windowState.start + windowState.size - 1] || "";
      label.textContent = total
        ? `${{startPeriod}} to ${{endPeriod}} | ${{windowState.size}} of ${{total}} period(s)`
        : "No periods";

      slider.oninput = () => {{
        windowState.start = Number(slider.value);
        renderChart();
      }};
      zoomIn.onclick = () => {{
        setTimeWindowSize(total, Math.max(2, Math.floor(windowState.size * 0.65)));
        renderChart();
      }};
      zoomOut.onclick = () => {{
        setTimeWindowSize(total, Math.min(total, Math.ceil(windowState.size / 0.65)));
        renderChart();
      }};
      reset.onclick = () => {{
        windowState.start = 0;
        windowState.size = total;
        renderChart();
      }};
    }}

    function renderFrequencyButtons() {{
      const container = document.getElementById("frequencyButtons");
      container.innerHTML = Object.keys(DATA.frequencies).map(freq => `
        <button type="button" data-frequency="${{esc(freq)}}" class="${{freq === state.frequency ? "active" : ""}}">
          ${{esc(DATA.frequencyLabels[freq] || freq)}}
        </button>
      `).join("");
      container.querySelectorAll("button").forEach(button => {{
        button.addEventListener("click", () => {{
          state.frequency = button.dataset.frequency;
          state.caseTypes = new Set(currentData().caseTypes);
          if (!currentData().details[state.detailTable]) {{
            state.detailTable = Object.keys(currentData().details)[0] || "";
          }}
          renderAll();
        }});
      }});
    }}

    function renderScaleButtons() {{
      document.querySelectorAll("#scaleButtons button").forEach(button => {{
        button.classList.toggle("active", button.dataset.scale === state.scale);
        button.onclick = () => {{
          state.scale = button.dataset.scale;
          renderAll();
        }};
      }});
    }}

    function renderMetricList() {{
      const container = document.getElementById("metricList");
      container.innerHTML = DATA.metrics.map(metric => `
        <label class="choice">
          <input type="checkbox" value="${{esc(metric)}}" ${{state.metrics.has(metric) ? "checked" : ""}}>
          <span>${{esc(metric)}}</span>
        </label>
      `).join("");
      container.querySelectorAll("input").forEach(input => {{
        input.addEventListener("change", () => {{
          if (input.checked) {{
            state.metrics.add(input.value);
          }} else {{
            state.metrics.delete(input.value);
          }}
          renderAll();
        }});
      }});
    }}

    function renderCaseList() {{
      const data = currentData();
      if (!state.caseTypes.size) state.caseTypes = new Set(data.caseTypes);
      const container = document.getElementById("caseList");
      container.innerHTML = data.caseTypes.map(caseType => `
        <label class="choice">
          <input type="checkbox" value="${{esc(caseType)}}" ${{state.caseTypes.has(caseType) ? "checked" : ""}}>
          <span>${{esc(caseType)}}</span>
        </label>
      `).join("");
      container.querySelectorAll("input").forEach(input => {{
        input.addEventListener("change", () => {{
          if (input.checked) {{
            state.caseTypes.add(input.value);
          }} else {{
            state.caseTypes.delete(input.value);
          }}
          renderAll();
        }});
      }});
    }}

    function bindToolbar() {{
      document.getElementById("selectCore").onclick = () => {{
        state.metrics = new Set(DATA.defaultMetrics);
        renderAll();
      }};
      document.getElementById("selectAllMetrics").onclick = () => {{
        state.metrics = new Set(DATA.metrics);
        renderAll();
      }};
      document.getElementById("clearMetrics").onclick = () => {{
        state.metrics = new Set();
        renderAll();
      }};
      document.getElementById("selectAllCases").onclick = () => {{
        state.caseTypes = new Set(currentData().caseTypes);
        renderAll();
      }};
      document.getElementById("detailSearch").addEventListener("input", event => {{
        state.detailSearch = event.target.value.toLowerCase();
        renderDetailTable();
      }});
    }}

    function renderMeta() {{
      const data = currentData();
      const control = data.control || {{}};
      const meta = document.getElementById("meta");
      meta.innerHTML = [
        ["Current Date", control["Current Date"]],
        ["Actual Cutoff", control["Actual Cutoff Date"]],
        ["Master File Cutoff", control["Master File Cutoff Date"]],
        ["Frequency", data.label]
      ].filter(item => item[1]).map(item => `<span>${{esc(item[0])}}: ${{esc(item[1])}}</span>`).join("");
      document.getElementById("subtitle").textContent = `${{data.label}} view across ${{data.caseTypes.length}} case type(s)`;
    }}

    function renderSummary() {{
      const rows = getSelectedRows();
      const selectedMetrics = DATA.metrics.filter(metric => state.metrics.has(metric));
      const {{ byMetricPeriod, periods }} = aggregateRows(rows);
      const latestPeriod = periods[periods.length - 1];
      const grid = document.getElementById("summaryGrid");

      if (!selectedMetrics.length || !periods.length) {{
        grid.innerHTML = `<div class="empty">No selected data</div>`;
        document.getElementById("summaryNote").textContent = "";
        return;
      }}

      grid.innerHTML = selectedMetrics.slice(0, 8).map(metric => {{
        const values = periods.map(period => byMetricPeriod.get(`${{metric}}||${{period}}`) || 0);
        const total = values.reduce((sum, value) => sum + value, 0);
        const latest = byMetricPeriod.get(`${{metric}}||${{latestPeriod}}`) || 0;
        const peak = Math.max(...values);
        const mainValue = metric.includes("FTE") ? peak : total;
        const mainLabel = metric.includes("FTE") ? "Peak" : "Total";
        return `
          <div class="tile">
            <div class="tile-title">${{esc(metric)}}</div>
            <div class="tile-value">${{formatNumber(mainValue, metric)}}</div>
            <div class="tile-sub">${{mainLabel}} | Latest ${{esc(latestPeriod)}}: ${{formatNumber(latest, metric)}}</div>
          </div>
        `;
      }}).join("");
      document.getElementById("summaryNote").textContent = `${{periods[0]}} to ${{latestPeriod}}`;
    }}

    function renderChart() {{
      const rows = getSelectedRows();
      const selectedMetrics = DATA.metrics.filter(metric => state.metrics.has(metric));
      const {{ byMetricPeriod, periods: allPeriods }} = aggregateRows(rows);
      renderTimeControls(allPeriods);
      const periods = getVisiblePeriods(allPeriods);
      const svg = document.getElementById("trendChart");
      const legend = document.getElementById("legend");

      if (!selectedMetrics.length || !allPeriods.length || !periods.length) {{
        svg.innerHTML = `<text x="550" y="190" text-anchor="middle" fill="#627174">No selected data</text>`;
        legend.innerHTML = "";
        document.getElementById("chartNote").textContent = "";
        return;
      }}

      const width = 1100;
      const height = 390;
      const margin = {{ top: 24, right: 32, bottom: 72, left: 78 }};
      const plotWidth = width - margin.left - margin.right;
      const plotHeight = height - margin.top - margin.bottom;

      const series = selectedMetrics.map(metric => {{
        const rawValues = periods.map(period => byMetricPeriod.get(`${{metric}}||${{period}}`) || 0);
        const maxValue = Math.max(...rawValues, 0);
        const values = state.scale === "indexed" && maxValue > 0
          ? rawValues.map(value => value / maxValue * 100)
          : rawValues;
        return {{ metric, rawValues, values }};
      }});

      const yMax = Math.max(...series.flatMap(item => item.values), 1);
      const yTop = yMax * 1.08;
      const xStep = periods.length <= 1 ? 0 : plotWidth / (periods.length - 1);
      const x = index => margin.left + index * xStep;
      const y = value => margin.top + plotHeight - (value / yTop) * plotHeight;
      const ticks = [0, 0.25, 0.5, 0.75, 1].map(value => value * yTop);

      let markup = "";
      markup += `<rect x="0" y="0" width="${{width}}" height="${{height}}" fill="#fff"></rect>`;
      for (const tick of ticks) {{
        const yy = y(tick);
        markup += `<line x1="${{margin.left}}" y1="${{yy}}" x2="${{width - margin.right}}" y2="${{yy}}" stroke="#e7ecee"></line>`;
        markup += `<text x="${{margin.left - 10}}" y="${{yy + 4}}" text-anchor="end" fill="#627174" font-size="11">${{formatNumber(tick, state.scale === "indexed" ? "" : "Volume")}}</text>`;
      }}
      markup += `<line x1="${{margin.left}}" y1="${{margin.top}}" x2="${{margin.left}}" y2="${{height - margin.bottom}}" stroke="#9aa8ab"></line>`;
      markup += `<line x1="${{margin.left}}" y1="${{height - margin.bottom}}" x2="${{width - margin.right}}" y2="${{height - margin.bottom}}" stroke="#9aa8ab"></line>`;

      const labelStep = Math.max(1, Math.ceil(periods.length / 12));
      periods.forEach((period, index) => {{
        if (index % labelStep === 0 || index === periods.length - 1) {{
          const xx = x(index);
          markup += `<text x="${{xx}}" y="${{height - margin.bottom + 24}}" text-anchor="middle" fill="#627174" font-size="11">${{esc(period)}}</text>`;
        }}
      }});

      for (const item of series) {{
        const color = metricColor(item.metric);
        const points = item.values.map((value, index) => `${{x(index)}},${{y(value)}}`).join(" ");
        markup += `<polyline fill="none" stroke="${{color}}" stroke-width="2.4" points="${{points}}"></polyline>`;
        if (periods.length <= 60) {{
          item.values.forEach((value, index) => {{
            const raw = item.rawValues[index];
            markup += `<circle cx="${{x(index)}}" cy="${{y(value)}}" r="3.2" fill="${{color}}" data-metric="${{esc(item.metric)}}" data-period="${{esc(periods[index])}}" data-value="${{raw}}"></circle>`;
          }});
        }}
      }}
      markup += `
        <g id="chartHoverGroup" style="display:none">
          <line id="chartHoverLine" x1="${{margin.left}}" y1="${{margin.top}}" x2="${{margin.left}}" y2="${{height - margin.bottom}}" stroke="#21353a" stroke-width="1.2" stroke-dasharray="4 4"></line>
          <text id="chartHoverLabel" x="${{margin.left}}" y="${{height - margin.bottom + 46}}" text-anchor="middle" fill="#21353a" font-size="11"></text>
        </g>
        <rect id="chartHoverOverlay" x="${{margin.left}}" y="${{margin.top}}" width="${{plotWidth}}" height="${{plotHeight}}" fill="transparent" style="cursor: crosshair" pointer-events="all"></rect>
      `;

      svg.innerHTML = markup;
      const periodInfoByPeriod = new Map(
        currentData().periods.map(periodInfo => [periodInfo.Period, periodInfo])
      );
      const hoverGroup = svg.querySelector("#chartHoverGroup");
      const hoverLine = svg.querySelector("#chartHoverLine");
      const hoverLabel = svg.querySelector("#chartHoverLabel");
      const overlay = svg.querySelector("#chartHoverOverlay");
      const tooltip = document.getElementById("pointTooltip");

      function moveTooltip(event) {{
        const padding = 14;
        const tooltipWidth = tooltip.offsetWidth || 260;
        const tooltipHeight = tooltip.offsetHeight || 160;
        let left = event.clientX + 14;
        let top = event.clientY + 14;
        if (left + tooltipWidth + padding > window.innerWidth) {{
          left = event.clientX - tooltipWidth - 14;
        }}
        if (top + tooltipHeight + padding > window.innerHeight) {{
          top = event.clientY - tooltipHeight - 14;
        }}
        tooltip.style.left = `${{Math.max(padding, left)}}px`;
        tooltip.style.top = `${{Math.max(padding, top)}}px`;
      }}

      function showChartTooltip(event) {{
        const svgBounds = svg.getBoundingClientRect();
        const viewX = (event.clientX - svgBounds.left) * width / svgBounds.width;
        const clampedX = Math.max(margin.left, Math.min(width - margin.right, viewX));
        const nearestIndex = periods.length <= 1
          ? 0
          : Math.max(0, Math.min(periods.length - 1, Math.round((clampedX - margin.left) / xStep)));
        const period = periods[nearestIndex];
        const periodInfo = periodInfoByPeriod.get(period) || {{}};
        const dateText = periodInfo["Period Start"] && periodInfo["Period End"]
          ? `${{periodInfo["Period Start"]}} to ${{periodInfo["Period End"]}}`
          : period;
        const xx = x(nearestIndex);

        hoverGroup.style.display = "block";
        hoverLine.setAttribute("x1", xx);
        hoverLine.setAttribute("x2", xx);
        hoverLabel.setAttribute("x", xx);
        hoverLabel.textContent = period;

        const valueRows = series.map(item => {{
          const rawValue = item.rawValues[nearestIndex] || 0;
          const scaledValue = item.values[nearestIndex] || 0;
          const scaleNote = state.scale === "indexed"
            ? ` <span class="muted">(${{formatNumber(scaledValue)}} indexed)</span>`
            : "";
          return `
            <div style="display:flex;align-items:center;gap:8px;justify-content:space-between;margin-top:5px">
              <span><span class="swatch" style="background:${{metricColor(item.metric)}}"></span> ${{esc(item.metric)}}</span>
              <b>${{formatNumber(rawValue, item.metric)}}${{scaleNote}}</b>
            </div>
          `;
        }}).join("");

        tooltip.style.display = "block";
        tooltip.innerHTML = `
          <b>${{esc(period)}}</b><br>
          <span class="muted">${{esc(dateText)}}</span>
          <div style="height:1px;background:#e3e9eb;margin:7px 0"></div>
          ${{valueRows}}
        `;
        moveTooltip(event);
      }}

      overlay.addEventListener("mousemove", showChartTooltip);
      overlay.addEventListener("mouseleave", () => {{
        hoverGroup.style.display = "none";
        tooltip.style.display = "none";
      }});

      legend.innerHTML = selectedMetrics.map(metric => `
        <span class="legend-item"><span class="swatch" style="background:${{metricColor(metric)}}"></span>${{esc(metric)}}</span>
      `).join("");
      document.getElementById("chartNote").textContent = state.scale === "indexed"
        ? "Indexed to each selected metric peak | Hover over the plot for date and values"
        : "Absolute values | Hover over the plot for date and values";
    }}

    function renderMatrix() {{
      const rows = getSelectedRows();
      const data = currentData();
      const periods = data.periods.map(period => period.Period);
      const selectedMetrics = DATA.metrics.filter(metric => state.metrics.has(metric));
      const selectedCases = data.caseTypes.filter(caseType => state.caseTypes.has(caseType));
      const map = new Map();

      for (const row of rows) {{
        const key = `${{row.Metric}}||${{row["Case Type"]}}||${{row.Period}}`;
        map.set(key, (map.get(key) || 0) + (Number(row.Volume) || 0));
      }}

      if (!selectedMetrics.length || !selectedCases.length || !periods.length) {{
        document.getElementById("matrixTable").innerHTML = `<div class="empty">No selected data</div>`;
        return;
      }}

      let html = `<table class="matrix-grid"><thead><tr><th class="sticky-col metric-col">Metric</th><th class="sticky-col case-col">Case Type</th>`;
      html += periods.map(period => `<th class="period-col">${{esc(period)}}</th>`).join("");
      html += "</tr></thead><tbody>";
      for (const metric of selectedMetrics) {{
        for (const caseType of selectedCases) {{
          html += `<tr><td class="sticky-col metric-col">${{esc(metric)}}</td><td class="sticky-col case-col">${{esc(caseType)}}</td>`;
          html += periods.map(period => {{
            const value = map.get(`${{metric}}||${{caseType}}||${{period}}`) || 0;
            return `<td class="period-col">${{formatNumber(value, metric)}}</td>`;
          }}).join("");
          html += "</tr>";
        }}
      }}
      html += "</tbody></table>";
      document.getElementById("matrixTable").innerHTML = html;
      document.getElementById("matrixNote").textContent = `${{selectedMetrics.length}} metric(s), ${{selectedCases.length}} case type(s)`;
    }}

    function renderDetailControls() {{
      const data = currentData();
      const select = document.getElementById("detailTableSelect");
      const tableNames = Object.keys(data.details);
      if (!tableNames.includes(state.detailTable)) {{
        state.detailTable = tableNames[0] || "";
      }}
      select.innerHTML = tableNames.map(name => `
        <option value="${{esc(name)}}" ${{name === state.detailTable ? "selected" : ""}}>${{esc(data.detailLabels[name] || name)}}</option>
      `).join("");
      select.onchange = () => {{
        state.detailTable = select.value;
        renderDetailTable();
      }};
      document.getElementById("detailSearch").value = state.detailSearch;
    }}

    function renderDetailTable() {{
      const data = currentData();
      const rows = data.details[state.detailTable] || [];
      const counts = data.detailCounts[state.detailTable] || {{ includedRows: rows.length, totalRows: rows.length }};
      const search = state.detailSearch;
      const filteredRows = search
        ? rows.filter(row => Object.values(row).some(value => String(value ?? "").toLowerCase().includes(search)))
        : rows;
      const maxRows = 500;
      const visibleRows = filteredRows.slice(0, maxRows);
      const columns = Array.from(new Set(visibleRows.flatMap(row => Object.keys(row))));
      const target = document.getElementById("detailTable");
      const loadedNote = counts.includedRows < counts.totalRows
        ? `${{counts.includedRows.toLocaleString()}} of ${{counts.totalRows.toLocaleString()}} rows loaded for browser performance`
        : `${{counts.totalRows.toLocaleString()}} rows loaded`;
      document.getElementById("detailNote").textContent = `${{filteredRows.length.toLocaleString()}} matching row(s), showing first ${{visibleRows.length.toLocaleString()}} | ${{loadedNote}}`;

      if (!visibleRows.length) {{
        target.innerHTML = `<div class="empty">No detail rows</div>`;
        return;
      }}

      let html = "<table><thead><tr>";
      html += columns.map(column => `<th>${{esc(column)}}</th>`).join("");
      html += "</tr></thead><tbody>";
      for (const row of visibleRows) {{
        html += "<tr>";
        html += columns.map(column => `<td>${{esc(formatNumber(row[column], column))}}</td>`).join("");
        html += "</tr>";
      }}
      html += "</tbody></table>";
      target.innerHTML = html;
    }}

    function renderAll() {{
      renderFrequencyButtons();
      renderScaleButtons();
      renderMetricList();
      renderCaseList();
      renderMeta();
      renderSummary();
      renderChart();
      renderMatrix();
      renderDetailControls();
      renderDetailTable();
    }}

    bindToolbar();
    state.caseTypes = new Set(currentData().caseTypes);
    renderAll();
  </script>
</body>
</html>
"""

    def write_html(self, output_path=None):
        if output_path is not None:
            self.output_path = output_path

        output_dir = os.path.dirname(self.output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        try:
            with open(self.output_path, "w", encoding="utf-8") as html_file:
                html_file.write(self.build_html())
        except PermissionError:
            base_path, extension = os.path.splitext(self.output_path)
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            self.output_path = f"{base_path}_{timestamp}{extension}"
            with open(self.output_path, "w", encoding="utf-8") as html_file:
                html_file.write(self.build_html())

        return self.output_path
