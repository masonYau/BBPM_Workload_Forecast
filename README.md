# Workload Forecast / 工作量预测

This project reads case tracker, beginning-of-workload plan, and completion probability inputs, then generates Excel and HTML workload forecast reports.

本项目读取案件跟踪表、期初/计划工作量输入和完成概率输入，并生成 Excel 与 HTML 工作量预测报表。

## Quick Start / 快速开始

Run from the repository root:

在项目根目录运行：

```powershell
C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\main.py
```

Use a custom config file:

使用自定义配置文件：

```powershell
C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\main.py --config .\config.json
```

Generated files / 生成文件：

| File / 文件 | Purpose / 用途 |
| --- | --- |
| `data/Workload_Forecast_Output.xlsx` | Excel forecast output. / Excel 预测结果。 |
| `data/Workload_Forecast_Visualization.html` | Interactive HTML visualization. / 交互式 HTML 可视化。 |
| `mi.log` | Local run log path configured by `outputs.log_path`. The same log messages are also printed to the terminal while running. / 由 `outputs.log_path` 配置的本地运行日志路径。运行时同一批日志也会打印到终端。 |

## Configuration File / 配置文件

All configuration is stored in `config.json`.

所有配置都在 `config.json` 中维护。

Path handling / 路径处理：

| Rule / 规则 | Explanation / 说明 |
| --- | --- |
| Input file paths are resolved from the current path first, then from `data/`. | 输入文件路径会先按当前路径查找，找不到再到 `data/` 目录查找。 |
| Tracker file path supports wildcards such as `*.xlsm`. If multiple files match, the latest modified file is used. | Tracker 文件支持 `*.xlsm` 等通配符；如匹配多个文件，会使用最后修改时间最新的文件。 |
| Output paths are used as written. Parent folders are created automatically. | 输出路径按配置写出，父目录会自动创建。 |

## Input Parameters / 输入参数

### `inputs.completion_percentage`

Completion probability distribution input. It is used to forecast when WIP and remaining planned starts will complete.

完成概率分布输入。用于预测 WIP 和剩余计划开始量在未来哪个周期完成。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `inputs.completion_percentage.file_path` | `Input_Completion_Percentage.xlsx` | Workbook containing completion probability curves. Changing it changes the source of forecast completion probabilities. / 完成概率曲线所在工作簿。修改后会改变预测完成量使用的概率来源。 |
| `inputs.completion_percentage.sheet_names` | `null` | `null` means read all sheets. A string or list restricts reading to selected sheets. Each valid sheet name becomes a case type, such as `PR` or `Trigger`. / `null` 表示读取全部 sheet；也可配置字符串或列表限制读取范围。每个有效 sheet 名会作为案件类型，例如 `PR` 或 `Trigger`。 |
| `inputs.completion_percentage.period_column` | `Month` | Column containing elapsed input-period numbers, not calendar dates. Values are converted to numeric period offsets. / 已流逝输入周期编号所在列，不是日历日期。该列会被转为数字周期偏移量。 |
| `inputs.completion_percentage.percentage_column` | `Percentage` | Column containing completion probability. If values are greater than 1, they are treated as percentages and divided by 100. / 完成概率列。如果数值大于 1，会按百分比处理并除以 100。 |
| `inputs.completion_percentage.frequency` | `M` | Frequency of the input probability curve. It is converted to the requested output frequency using `run.frequency_days`. Wrong frequency will shift forecast completion timing. / 输入概率曲线的周期颗粒度。程序会用 `run.frequency_days` 转换到输出颗粒度；配置错误会导致预测完成时间偏移。 |

### `inputs.bow_volume`

Beginning-of-workload or planned start volume input. It is used as the plan baseline and to calculate remaining planned starts.

期初/计划开始量输入。用于作为计划基线，并计算剩余计划开始量。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `inputs.bow_volume.file_path` | `Input_BoW_Volume.xlsx` | Workbook containing planned start volumes. Changing it changes the plan baseline. / 计划开始量所在工作簿。修改后会改变计划基线。 |
| `inputs.bow_volume.sheet_names` | `null` | `null` means read all sheets. A string or list restricts reading to selected sheets. Each valid sheet name becomes a case type. / `null` 表示读取全部 sheet；也可配置字符串或列表限制读取范围。每个有效 sheet 名会作为案件类型。 |
| `inputs.bow_volume.period_column` | `Month` | Calendar period/date column. Values are parsed as dates and converted to periods using `inputs.bow_volume.frequency`. / 日历周期或日期列。程序会按日期解析，再根据 `inputs.bow_volume.frequency` 转换为周期。 |
| `inputs.bow_volume.volume_column` | `Volume` | Planned volume column. Values are converted to numbers and summed by case type and input period. / 计划量列。程序会转为数值，并按案件类型和输入周期汇总。 |
| `inputs.bow_volume.frequency` | `M` | Frequency of the BoW input. It controls how monthly/weekly/daily allocation is converted to output periods. Wrong frequency will distort planned and remaining start volumes. / BoW 输入颗粒度。它决定计划量如何分摊到月/周/日输出周期；配置错误会扭曲计划开始量和剩余计划开始量。 |

### `inputs.tracker`

Case tracker input. It is used to calculate received starts, actual starts, actual completions, WIP starts, and cutoff dates.

案件跟踪表输入。用于计算已收到开始量、实际开始量、实际完成量、WIP 开始量和截断日期。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `inputs.tracker.file_path` | `BBPM Case tracker(Updated) - *.xlsm` | Tracker workbook path or wildcard pattern. If multiple files match, the latest modified file is selected. / Tracker 工作簿路径或通配符。如匹配多个文件，会选择最后修改时间最新的文件。 |
| `inputs.tracker.sheet_name` | `Master File` | Sheet read from the tracker workbook. It must contain the required columns such as `CIN`, `Original T0`, `Review Type`, `Task Status`, and `Approval/Cancel Date`. / Tracker 中读取的 sheet。必须包含 `CIN`、`Original T0`、`Review Type`、`Task Status`、`Approval/Cancel Date` 等关键列。 |
| `inputs.tracker.skiprows` | `1` | Number of rows skipped before reading headers. Changing it changes which row pandas treats as the header row. / 读取表头前跳过的行数。修改后会影响哪一行被识别为列名。 |

## Run Parameters / 运行参数

### `run`

Run-level parameters control output frequency selection and frequency conversion.

运行级参数控制输出颗粒度选择和不同颗粒度之间的换算。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `run.default_forecast_frequency` | `M` | Default frequency for the Excel report and the first processor built by `main.py`. Values should exist in `run.frequency_days`. / Excel 报表默认输出颗粒度，也是 `main.py` 首先构建的处理器颗粒度。取值应存在于 `run.frequency_days`。 |
| `run.visualization_frequencies` | `["M", "W", "D"]` | Compatibility field. The current default HTML generation ignores this list and includes all frequencies defined in `run.frequency_days`, ordered as `M`, `W`, `D` when present and then any extra configured keys. If calling `run_workload_visualization(..., frequencies=...)` directly, that explicit function argument still controls HTML frequencies. / 兼容字段。当前默认 HTML 生成会忽略该列表，并包含 `run.frequency_days` 中定义的全部颗粒度；如存在 `M`、`W`、`D` 会按该顺序展示，之后追加其他配置频率。如果直接调用 `run_workload_visualization(..., frequencies=...)`，显式传入的函数参数仍会控制 HTML 颗粒度。 |
| `run.frequency_days` | `{"D": 1, "W": 7, "M": 30}` | Conversion table from frequency code to approximate day count. It drives completion-distribution conversion, WBH timing conversion, and the list of default HTML frequencies. Changing values changes timing allocation and may materially change forecast results. / 频率代码到近似天数的换算表。它影响完成概率转换、WBH 时间点转换，也决定默认 HTML 包含哪些颗粒度。修改数值会改变时间分摊，可能显著影响预测结果。 |

Frequency code details / 频率代码说明：

| Code / 代码 | Days / 天数 | Meaning / 含义 | Impact / 影响 |
| --- | ---: | --- | --- |
| `D` | `1` | Daily. / 日。 | Produces daily output periods and the most detailed HTML view. / 生成日粒度输出和最细 HTML 视图。 |
| `W` | `7` | Weekly. / 周。 | Produces weekly periods. WBH day rules are converted by integer division. / 生成周粒度周期。WBH 天数规则会通过整数除法转换为周数。 |
| `M` | `30` | Monthly approximation. / 月近似值。 | Produces monthly periods. Used as the default Excel frequency. / 生成月粒度周期，并作为默认 Excel 颗粒度。 |

## Business Rule Parameters / 业务规则参数

### Status parameters / 状态参数

Statuses from the tracker are first normalized through `business_rules.status_mapping`. Downstream rules then use the normalized status values.

Tracker 中的原始状态会先通过 `business_rules.status_mapping` 标准化，后续规则都使用标准化后的状态值。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `business_rules.completed_source_statuses` | `["E2E Completed", "ReCompleted - WBH"]` | Legacy/reserved parameter. It is loaded into the processor, but current calculations use `status_mapping` plus `actual_completion_statuses` for completion logic. / 历史兼容/预留参数。当前会被读取到处理器中，但实际完成逻辑使用 `status_mapping` 和 `actual_completion_statuses`。 |
| `business_rules.status_mapping` | See table below. / 见下表。 | Maps raw tracker statuses to normalized statuses: `Completed`, `WBH`, `Cancelled/Closed`, `WIP`, `Pending QC/BA`, or `Not Initiated`. Unknown statuses become missing and are excluded from status-dependent calculations. / 将 Tracker 原始状态映射为标准状态：`Completed`、`WBH`、`Cancelled/Closed`、`WIP`、`Pending QC/BA` 或 `Not Initiated`。未知状态会变为空值，并在依赖状态的计算中被排除。 |
| `business_rules.actual_start_statuses` | `["Completed", "WBH", "Cancelled/Closed", "WIP", "Pending QC/BA"]` | Normalized statuses counted as actual starts. Adding a status increases actual start volume; removing one reduces it. / 会被计入实际开始量的标准状态。增加状态会提高实际开始量，移除状态会降低实际开始量。 |
| `business_rules.open_start_statuses` | `["WBH", "WIP", "Not Initiated", "Pending QC/BA"]` | Normalized statuses treated as open/WIP received starts. These feed forecast completion and WBH action calculations. / 被视为开放中/WIP 已收到开始量的标准状态。这些量会进入预测完成和 WBH 动作预测。 |
| `business_rules.pending_qc_ba_statuses` | `["Pending QC/BA"]` | Normalized statuses counted in the standalone `Pending QC/BA Volume` metric. / 会单独计入 `Pending QC/BA Volume` 指标的标准状态。 |
| `business_rules.actual_completion_statuses` | `["Completed"]` | Normalized statuses counted as actual completions using `Approval/Cancel Date`. Adding statuses increases actual completion volume. / 会按 `Approval/Cancel Date` 计入实际完成量的标准状态。增加状态会提高实际完成量。 |

Current `status_mapping` entries / 当前 `status_mapping` 明细：

| Raw tracker status / Tracker 原始状态 | Normalized status / 标准状态 | Impact / 影响 |
| --- | --- | --- |
| `CSEM` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `Cancelled - 2nd AC Opening under NTB` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - AC re-opened under NTB` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - All AC Closed` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `Cancelled - CDAB Code` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - China Notch-Down` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - KPMG Managed` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - RM Managed` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cancelled - trigger reason can be discounted` | `Cancelled/Closed` | Counted as actual start, not actual completion. / 计入实际开始，不计入实际完成。 |
| `Cat C CSEM - CDD deficiencies (Geographical risk)` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `Cat C CSEM - CDD deficienes (Geographical risk)` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `Cat C CSEM - Tax Risk Appetite` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `COMPLETED` | `WIP` | Currently treated as WIP, so it is counted as actual start and WIP received start. / 当前按 WIP 处理，因此计入实际开始和 WIP 已收到开始。 |
| `E2E Completed` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `FAILED` | `WIP` | Counted as actual start and WIP received start. / 计入实际开始和 WIP 已收到开始。 |
| `IN_PROGRESS` | `WIP` | Counted as actual start and WIP received start. / 计入实际开始和 WIP 已收到开始。 |
| `Not Loading` | `Not Initiated` | Counted as WIP received start, not actual start. / 计入 WIP 已收到开始，不计入实际开始。 |
| `Pending BA Approval` | `Pending QC/BA` | Counted as actual start, WIP received start, and `Pending QC/BA Volume`. / 计入实际开始、WIP 已收到开始和 `Pending QC/BA Volume`。 |
| `Pending Data Capture` | `WIP` | Counted as actual start and WIP received start. / 计入实际开始和 WIP 已收到开始。 |
| `Pending QC` | `Pending QC/BA` | Counted as actual start, WIP received start, and `Pending QC/BA Volume`. / 计入实际开始、WIP 已收到开始和 `Pending QC/BA Volume`。 |
| `WBH Imposed` | `WBH` | Counted as actual start and WIP received start. / 计入实际开始和 WIP 已收到开始。 |
| `ReCompleted - WBH` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |
| `Completed - WBH with KYCH over 30 days` | `Completed` | Counted as actual start and actual completion. / 计入实际开始和实际完成。 |

### WBH action parameters / WBH 动作参数

These parameters set the day thresholds for forecasting WBH letter and call workload. The code converts days to output periods using integer division: `days // run.frequency_days[output_frequency]`.

这些参数设置预测 WBH letter 和 call 工作量的天数阈值。代码使用整数除法把天数转换为输出周期：`days // run.frequency_days[output_frequency]`。

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `business_rules.wbh_letter_days.PR` | `90` | PR cases are forecast to need a WBH letter at T+90 days if still uncompleted. / PR 案件如果到 T+90 天仍未完成，则预测需要 WBH letter。 |
| `business_rules.wbh_letter_days.Trigger` | `60` | Trigger cases are forecast to need a WBH letter at T+60 days if still uncompleted. / Trigger 案件如果到 T+60 天仍未完成，则预测需要 WBH letter。 |
| `business_rules.wbh_call_days.PR` | `95` | PR cases are forecast to need a WBH call at T+95 days if still uncompleted. / PR 案件如果到 T+95 天仍未完成，则预测需要 WBH call。 |
| `business_rules.wbh_call_days.Trigger` | `65` | Trigger cases are forecast to need a WBH call at T+65 days if still uncompleted. / Trigger 案件如果到 T+65 天仍未完成，则预测需要 WBH call。 |

### Demand and probability parameters / 需求和概率参数

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `business_rules.completion_upt` | `4.5` | Unit processing time for completion volume. Used in demand hours: `Completion Volume * Completion UPT`. Higher values increase `Demand FTE`. / 完成量的单位处理时间。用于需求工时：`Completion Volume * Completion UPT`。数值越高，`Demand FTE` 越高。 |
| `business_rules.init_upt` | `0.5` | Unit processing time for init/start volume. Used in demand hours: `Init Volume * Init UPT`. Higher values increase `Demand FTE`. / 初始化/开始量的单位处理时间。用于需求工时：`Init Volume * Init UPT`。数值越高，`Demand FTE` 越高。 |
| `business_rules.working_hour` | `129` | Available working hours per FTE per period. Demand FTE is calculated as `(Completion Volume * completion_upt + Init Volume * init_upt) / working_hour`. Lower values increase `Demand FTE`; zero or missing suppresses FTE calculation. / 每个 FTE 每周期可用工作小时。`Demand FTE` 公式为 `(Completion Volume * completion_upt + Init Volume * init_upt) / working_hour`。数值越低，`Demand FTE` 越高；为 0 或缺失时不会正常计算 FTE。 |
| `business_rules.read_completion_percentage_from_input` | `true` | `true` reads completion probability from `Input_Completion_Percentage.xlsx`. `false` calculates a distribution from tracker history. Changing this changes the source of forecast completion and WBH survival probabilities. / `true` 表示从 `Input_Completion_Percentage.xlsx` 读取完成概率；`false` 表示从 Tracker 历史数据计算分布。修改该参数会改变预测完成和 WBH 未完成概率来源。 |

## Output Parameters / 输出参数

| Parameter / 参数 | Current value / 当前值 | Function and impact / 作用和影响 |
| --- | --- | --- |
| `outputs.excel_path` | `data/Workload_Forecast_Output.xlsx` | Default Excel output path. If the file is locked, the program writes a timestamped fallback file. / 默认 Excel 输出路径。如果文件被占用，程序会写出带时间戳的备用文件。 |
| `outputs.html_path` | `data/Workload_Forecast_Visualization.html` | Default HTML visualization output path. If the file is locked, the program writes a timestamped fallback file. / 默认 HTML 可视化输出路径。如果文件被占用，程序会写出带时间戳的备用文件。 |
| `outputs.log_path` | `./mi.log` | Log output path. It controls the file log location; the same log messages are also printed to the terminal. Relative paths are resolved from the current working directory. / 日志输出路径。它控制文件日志位置；同一批日志也会打印到终端。相对路径按当前运行目录解析。 |

## Report Contents / 报表内容

The Excel output writes a control sheet, wide and long workload tables, metric definitions, source/plan tables, start reconciliation, WIP and Pending QC/BA detail, completion detail, WBH action detail, and monthly calculation audit sheets.

Excel 输出会写入控制页、宽表和长表工作量、指标定义、源数据/计划表、开始量对账、WIP 与 Pending QC/BA 明细、完成量明细、WBH 动作明细，以及月度计算审计页。

| Output area / 输出区域 | Main sheets or views / 主要 sheet 或视图 |
| --- | --- |
| Control and definitions / 控制和定义 | `00_Control`, `03_Metric_Definitions` |
| Workload tables / 工作量表 | `01_Workload_Wide`, `02_Workload_Long` |
| Start pipeline / 开始量链路 | `10_Input_BoW`, `11_Planned_Start`, `12_Start_Reconciliation`, `13_WIP_Received_Start`, `14_Pending_QC_BA` |
| Completion and WBH / 完成量和 WBH | `20_Completion`, `21_Completion_Distribution`, `30_WBH_Letter`, `31_WBH_Call` |
| Monthly calculation audit / 月度计算审计 | `40_Calc_Forecast_Completion`, `41_Calc_WBH_Letter`, `42_Calc_WBH_Call`, `43_Calc_Demand_FTE` |
| HTML visualization / HTML 可视化 | Frequency switcher, metric filters, case-type filters, summary cards, trend chart, metric matrix, and searchable detail tables. / 频率切换、指标筛选、案件类型筛选、汇总卡片、趋势图、指标矩阵和可搜索明细表。 |

## Python Entry Parameters / Python 入口函数参数

These parameters are used when calling functions from Python instead of running `main.py` directly.

这些参数用于从 Python 中直接调用函数，而不是直接运行 `main.py`。

| Function / 函数 | Parameter / 参数 | Default / 默认值 | Function and impact / 作用和影响 |
| --- | --- | --- | --- |
| `setup_logging` | `log_path` | `outputs.log_path` | Overrides the configured log output file path when provided. / 传入时覆盖配置中的日志输出路径。 |
| `setup_logging` | `level` | `logging.INFO` | Sets minimum log level for both file and terminal logs. Lower levels such as `DEBUG` can produce more detail; higher levels such as `WARNING` reduce routine progress logs. / 同时设置文件日志和终端日志的最低级别。较低级别如 `DEBUG` 会输出更多细节，较高级别如 `WARNING` 会减少常规进度日志。 |
| `build_workload_processor` | `current_date` | Required / 必填 | Sets the current/cutoff date used by one processor. / 设置单个处理器使用的当前日期和截断日期。 |
| `build_workload_processor` | `frequency` | Required / 必填 | Sets the output frequency for one processor. It must exist in `run.frequency_days`. / 设置单个处理器的输出颗粒度，必须存在于 `run.frequency_days`。 |
| `build_workload_processor` | `config` | `load_config()` inside `DataProcessor` | Optional in-memory config override. Passing the same config avoids reloading `config.json`. / 可选的内存配置覆盖。传入同一份配置可避免重复读取 `config.json`。 |
| `main` | `config_path` | `None` | Optional path to a config file. When omitted, the default repository `config.json` is used. The CLI supports both `python main.py path\to\config.json` and `python main.py --config path\to\config.json`. / 可选配置文件路径。省略时使用仓库默认 `config.json`。命令行同时支持 `python main.py path\to\config.json` 和 `python main.py --config path\to\config.json`。 |
| `run_workload_forecast` | `current_date` | `datetime.now()` | Sets the cutoff/current date for calculations. / 设置计算使用的当前日期和截断日期。 |
| `run_workload_forecast` | `frequency` | `run.default_forecast_frequency` | Sets the Excel forecast frequency. / 设置 Excel 预测颗粒度。 |
| `run_workload_forecast` | `output_path` | `outputs.excel_path` | Overrides the Excel output path for that call. / 覆盖本次调用的 Excel 输出路径。 |
| `run_workload_forecast` | `config_path` | `None` | Optional config file path for that call. / 本次调用可选的配置文件路径。 |
| `run_workload_visualization` | `current_date` | `datetime.now()` | Sets the cutoff/current date for all HTML frequency calculations. / 设置 HTML 各颗粒度计算使用的当前日期和截断日期。 |
| `run_workload_visualization` | `frequencies` | All keys from `run.frequency_days` | Explicitly controls HTML frequencies when provided. If omitted, all configured frequencies are included. / 显式传入时控制 HTML 包含哪些颗粒度；省略时包含 `run.frequency_days` 中全部颗粒度。 |
| `run_workload_visualization` | `output_path` | `outputs.html_path` | Overrides the HTML output path for that call. / 覆盖本次调用的 HTML 输出路径。 |
| `run_workload_visualization` | `config_path` | `None` | Optional config file path for that call. / 本次调用可选的配置文件路径。 |
| `WorkloadHtmlVisualizer` | `max_detail_rows` | `5000` | Caps rows embedded for each detail table in the HTML payload. Higher values show more detail but increase HTML size and browser load time. / 限制 HTML 中每张明细表嵌入的最大行数。数值越高，明细越完整，但 HTML 文件更大、浏览器加载更慢。 |

## Source Archive Export / 源文件归档导出

Run `copy_py_to_md.py` before committing. It exports both Python and JSON source files into the `copy/` folder as Markdown files with matching code fences.

提交前运行 `copy_py_to_md.py`。它会把 Python 和 JSON 源文件都导出到 `copy/` 目录，生成带对应代码块语言标记的 Markdown 归档。

## Key Metric Effects / 关键指标影响关系

| Metric / 指标 | Main drivers / 主要驱动参数 |
| --- | --- |
| `Planned Start Volume` | `inputs.bow_volume.*`, `run.frequency_days` |
| `Received Start Volume` | `inputs.tracker.*`, `Original T0`, cutoff date |
| `WIP Received Start Volume` | `inputs.tracker.*`, `business_rules.status_mapping`, `business_rules.open_start_statuses`, BoW cutoff date |
| `Pending QC/BA Volume` | `inputs.tracker.*`, `business_rules.status_mapping`, `business_rules.pending_qc_ba_statuses`, BoW cutoff date |
| `Remaining Planned Start Volume` | `inputs.bow_volume.*`, tracker received starts, actual cutoff date, BoW cutoff date |
| `Actual Start Volume` | `inputs.tracker.*`, `business_rules.status_mapping`, `business_rules.actual_start_statuses` |
| `Actual Completion Volume` | `inputs.tracker.*`, `business_rules.status_mapping`, `business_rules.actual_completion_statuses` |
| `Forecast Completion Volume` | `business_rules.read_completion_percentage_from_input`, `inputs.completion_percentage.*`, remaining/WIP starts |
| `Completion Volume` | actual completion volume plus forecast completion volume |
| `Init Volume` | received start volume plus remaining planned start volume |
| `Forecast WBH Letter Volume` | `business_rules.wbh_letter_days.*`, completion probability, remaining/WIP starts |
| `Forecast WBH Call Volume` | `business_rules.wbh_call_days.*`, completion probability, remaining/WIP starts |
| `Demand FTE` | `business_rules.completion_upt`, `business_rules.init_upt`, `business_rules.working_hour`, completion/init volumes |
