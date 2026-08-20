# data_processor.py

```python
import pandas as pd

import logging
import numpy as np
import math
import os
import time
try:
    from .headers import MasterHeader as mh
    from .headers import BowHeader as bh
    from .headers import RamHeader as rh
    from .config_loader import load_config
except ImportError:
    from headers import MasterHeader as mh
    from headers import BowHeader as bh
    from headers import RamHeader as rh
    from config_loader import load_config

logger = logging.getLogger(__name__)


class DataProcessor:

    def __init__(self, current_date: pd.Timestamp, frequency: str, config=None):

        self.config = config or load_config()
        input_config = self.config["inputs"]
        run_config = self.config["run"]
        business_rules = self.config["business_rules"]
        self.master_df = pd.DataFrame()
        self.comple_status = set(business_rules["completed_source_statuses"])
        self.wbh_status = dict(business_rules["status_mapping"])
        self.current_date = current_date
        self.frequency = frequency
        self.frequency_days = dict(run_config["frequency_days"])
        self.completion_distribution = pd.Series()
        self.input_bow_volume = pd.Series()
        self.bow_volume = pd.Series()
        self.received_volume = pd.Series()
        self.actual_start_volume = pd.Series()
        self.open_received_start_volume = pd.Series()
        self.pending_qc_ba_volume = pd.Series()
        self.forecast_completion_volume = pd.Series()
        self.actual_completion_volume = pd.Series()
        self.forecast_wbh_letter_volume = pd.Series()
        self.forecast_wbh_call_volume = pd.Series()
        self.completion_volume = pd.DataFrame()
        self.wbh_letter_volume = pd.DataFrame()
        self.wbh_call_volume = pd.DataFrame()
        self.workload_volume = pd.DataFrame()
        self.remaining_bow_volume = pd.DataFrame()
        self.actual_cutoff_date = pd.NaT
        self.remaining_bow_cutoff_date = pd.NaT
        self.actual_start_status = set(business_rules["actual_start_statuses"])
        self.open_start_status = set(business_rules["open_start_statuses"])
        self.pending_qc_ba_status = set(business_rules.get("pending_qc_ba_statuses", ["Pending QC/BA"]))
        self.actual_completion_status = set(business_rules["actual_completion_statuses"])
        self.wbh_letter_days = dict(business_rules["wbh_letter_days"])
        self.wbh_call_days = dict(business_rules["wbh_call_days"])
        self.completion_upt = business_rules["completion_upt"]
        self.init_upt = business_rules["init_upt"]
        self.working_hour = business_rules["working_hour"]
        self.output_metric_definitions = [
            {
                "Category": "Start Pipeline",
                "Subcategory": "Plan",
                "Metric": "Planned Start Volume",
                "Display Name": "Planned Start Volume",
                "Source": "Input_BoW_Volume.xlsx",
                "Logic": "Monthly input BoW volume allocated to output periods by calendar-day overlap.",
                "Cutoff": "N/A",
            },
            {
                "Category": "Start Pipeline",
                "Subcategory": "Master File Received",
                "Metric": "Received Start Volume",
                "Display Name": "Received Start Volume",
                "Source": "Master File",
                "Logic": "All received cases through the BoW cutoff aggregated by Original T0 and output period.",
                "Cutoff": "BoW cutoff",
            },
            {
                "Category": "Start Pipeline",
                "Subcategory": "Master File WIP",
                "Metric": "WIP Received Start Volume",
                "Display Name": "WIP Received Start Volume",
                "Source": "Master File",
                "Logic": "Received cases still open after excluding completed and cancelled/closed status.",
                "Cutoff": "BoW cutoff",
            },
            {
                "Category": "Start Pipeline",
                "Subcategory": "Master File Pending",
                "Metric": "Pending QC/BA Volume",
                "Display Name": "Pending QC/BA Volume",
                "Source": "Master File",
                "Logic": "Received cases mapped to Pending QC/BA, aggregated by Original T0 and output period.",
                "Cutoff": "BoW cutoff",
            },
            {
                "Category": "Start Pipeline",
                "Subcategory": "Remaining Plan",
                "Metric": "Remaining Planned Start Volume",
                "Display Name": "Remaining Planned Start Volume",
                "Source": "Input_BoW_Volume.xlsx + Master File",
                "Logic": (
                    "Monthly: allocate input BoW from the actual cutoff. Weekly/daily: allocate "
                    "max(input BoW - received through actual cutoff at input frequency, 0). "
                    "Remaining per output period is max(output plan - received through BoW cutoff, 0)."
                ),
                "Cutoff": "Actual cutoff for plan basis/allocation; BoW cutoff for output-period received volume.",
            },
            {
                "Category": "Start Pipeline",
                "Subcategory": "Actual",
                "Metric": "Actual Start Volume",
                "Display Name": "Actual Start Volume",
                "Source": "Master File",
                "Logic": "Cases with status in Completed, WBH, Cancelled/Closed, WIP, Pending QC/BA aggregated by Original T0.",
                "Cutoff": "Current date",
            },
            {
                "Category": "Completion",
                "Subcategory": "Actual",
                "Metric": "Actual Completion Volume",
                "Display Name": "Actual Completion Volume",
                "Source": "Master File",
                "Logic": "Completed cases aggregated by Approval/Cancel Date.",
                "Cutoff": "Current date",
            },
            {
                "Category": "Completion",
                "Subcategory": "Forecast",
                "Metric": "Forecast Completion Volume",
                "Display Name": "Forecast Completion Volume",
                "Source": "WIP received starts + remaining planned starts + completion distribution",
                "Logic": "WIP received and remaining planned starts multiplied by case-type completion probability distribution.",
                "Cutoff": "BoW cutoff for WIP received; remaining plan follows actual/BoW cutoff rules.",
            },
            {
                "Category": "Demand",
                "Subcategory": "Input Volume",
                "Metric": "Completion Volume",
                "Display Name": "Completion Volume",
                "Source": "Actual completion + forecast completion",
                "Logic": "Actual Completion Volume plus Forecast Completion Volume.",
                "Cutoff": "Current date for actual completion; forecast follows source start cutoff rules.",
            },
            {
                "Category": "Demand",
                "Subcategory": "Input Volume",
                "Metric": "Init Volume",
                "Display Name": "Init Volume",
                "Source": "Received start + remaining planned start",
                "Logic": "Received Start Volume plus Remaining Planned Start Volume.",
                "Cutoff": "BoW cutoff for received starts; remaining plan follows actual/BoW cutoff rules.",
            },
            {
                "Category": "Demand",
                "Subcategory": "FTE",
                "Metric": "Demand FTE",
                "Display Name": "Demand FTE",
                "Source": "Completion Volume, Init Volume, UPT and working-hour assumptions",
                "Logic": "(Completion Volume * Completion UPT + Init Volume * Init UPT) / Working Hour.",
                "Cutoff": "Monthly report metric.",
            },
            {
                "Category": "WBH Action",
                "Subcategory": "Letter",
                "Metric": "Forecast WBH Letter Volume",
                "Display Name": "Forecast WBH Letter Volume",
                "Source": "WIP received starts + remaining planned starts + completion distribution",
                "Logic": "Starts still uncompleted at T+90 for PR or T+60 for Trigger.",
                "Cutoff": "BoW cutoff for WIP received; remaining plan follows actual/BoW cutoff rules.",
            },
            {
                "Category": "WBH Action",
                "Subcategory": "Call",
                "Metric": "Forecast WBH Call Volume",
                "Display Name": "Forecast WBH Call Volume",
                "Source": "WIP received starts + remaining planned starts + completion distribution",
                "Logic": "Starts still uncompleted at T+95 for PR or T+65 for Trigger.",
                "Cutoff": "BoW cutoff for WIP received; remaining plan follows actual/BoW cutoff rules.",
            },
        ]

        self.input_completion_percentage_config = dict(input_config["completion_percentage"])
        self.input_bow_volume_config = dict(input_config["bow_volume"])
        self.input_tracker_config = dict(input_config["tracker"])
        self.read_completion_percentage_from_input = business_rules["read_completion_percentage_from_input"]
        self.logged_summary_keys = set()

    def period_summary(self, data):
        if data is None or len(data) == 0:
            return 0, None, None

        periods = None
        if isinstance(data.index, pd.MultiIndex):
            for level_name in ("Period", "Start Period", "Input Period"):
                if level_name in data.index.names:
                    periods = data.index.get_level_values(level_name)
                    break
        elif isinstance(data.index, pd.PeriodIndex):
            periods = data.index
        elif isinstance(data, pd.DataFrame):
            for column in ("Period", "Start Period", "Input Period"):
                if column in data.columns:
                    periods = data[column]
                    break

        if periods is None:
            return 0, None, None

        period_values = sorted({str(period) for period in periods if pd.notna(period)})
        if not period_values:
            return 0, None, None
        return len(period_values), period_values[0], period_values[-1]

    def case_type_count(self, data):
        if data is None or len(data) == 0:
            return 0
        if isinstance(data.index, pd.MultiIndex) and "Case Type" in data.index.names:
            return data.index.get_level_values("Case Type").nunique()
        if isinstance(data, pd.DataFrame) and "Case Type" in data.columns:
            return data["Case Type"].nunique()
        return 0

    def numeric_total(self, data):
        if data is None or len(data) == 0:
            return 0.0
        if isinstance(data, pd.Series):
            return float(pd.to_numeric(data, errors="coerce").fillna(0).sum())

        numeric_data = data.select_dtypes(include=[np.number])
        if numeric_data.empty:
            return 0.0
        return float(numeric_data.sum().sum())

    def log_series_summary(self, label, series):
        period_count, period_start, period_end = self.period_summary(series)
        total = round(self.numeric_total(series), 2)
        summary_key = (
            label,
            self.frequency,
            0 if series is None else len(series),
            period_count,
            period_start,
            period_end,
            self.case_type_count(series),
            total,
        )
        if summary_key in self.logged_summary_keys:
            return
        self.logged_summary_keys.add(summary_key)

        logger.info(
            "%s summary | frequency=%s rows=%d periods=%d period_start=%s period_end=%s case_types=%d total=%.2f",
            label,
            self.frequency,
            0 if series is None else len(series),
            period_count,
            period_start,
            period_end,
            self.case_type_count(series),
            total,
        )

    def log_workload_summary(self):
        if self.workload_volume.empty:
            logger.warning("Workload summary | frequency=%s rows=0", self.frequency)
            return

        period_count, period_start, period_end = self.period_summary(self.workload_volume)
        metrics = [
            "Planned Start Volume",
            "Received Start Volume",
            "Pending QC/BA Volume",
            "Actual Completion Volume",
            "Forecast Completion Volume",
            "Demand FTE",
        ]
        totals = {
            metric: round(float(pd.to_numeric(self.workload_volume[metric], errors="coerce").fillna(0).sum()), 2)
            for metric in metrics
            if metric in self.workload_volume.columns
        }
        summary_key = (
            "Workload summary",
            self.frequency,
            len(self.workload_volume),
            period_count,
            period_start,
            period_end,
            self.case_type_count(self.workload_volume),
            tuple(sorted(totals.items())),
        )
        if summary_key in self.logged_summary_keys:
            return
        self.logged_summary_keys.add(summary_key)

        logger.info(
            "Workload summary | frequency=%s rows=%d periods=%d period_start=%s period_end=%s case_types=%d totals=%s",
            self.frequency,
            len(self.workload_volume),
            period_count,
            period_start,
            period_end,
            self.case_type_count(self.workload_volume),
            totals,
        )

    def read_data(self):
        from glob import glob

        tracker_path_pattern = self.resolve_input_file_pattern(self.input_tracker_config["file_path"])
        sheet_name = self.input_tracker_config["sheet_name"]
        skiprows = self.input_tracker_config.get("skiprows")

        matches = glob(tracker_path_pattern)
        if not matches:
            logger.error(
                "Case tracker not found | frequency=%s pattern=%s",
                self.frequency,
                tracker_path_pattern,
            )
            raise FileNotFoundError(f"No headcount files matched: {tracker_path_pattern}")

        tracker_path = max(matches, key=os.path.getmtime)
        logger.info(
            "Reading case tracker | frequency=%s file=%s sheet=%s skiprows=%s matched_files=%d",
            self.frequency,
            tracker_path,
            sheet_name,
            skiprows,
            len(matches),
        )
        master_df = pd.read_excel(tracker_path, sheet_name=sheet_name, skiprows=skiprows)
        required_columns = [mh.CIN, mh.OriginalT0, mh.ReviewType, mh.TaskStatus, mh.ApprovalCancelDate]
        missing_columns = [column for column in required_columns if column not in master_df.columns]
        if missing_columns:
            logger.error(
                "Case tracker missing required columns | frequency=%s missing_columns=%s",
                self.frequency,
                missing_columns,
            )

        master_df[mh.CIN] = master_df[mh.CIN].astype(str).apply(lambda x: x.lstrip('0'))
        master_df[mh.OriginalT0] = pd.to_datetime(master_df[mh.OriginalT0])
        master_df[mh.ReviewType] = master_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in x else "Trigger")
        unknown_statuses = master_df.loc[
            master_df[mh.TaskStatus].map(self.wbh_status).isna(),
            mh.TaskStatus,
        ].dropna().astype(str).unique()
        if len(unknown_statuses) > 0:
            logger.warning(
                "Unknown task statuses found | frequency=%s count=%d statuses=%s",
                self.frequency,
                len(unknown_statuses),
                sorted(unknown_statuses),
            )

        original_t0 = pd.to_datetime(master_df[mh.OriginalT0], errors="coerce").dropna()
        logger.info(
            "Case tracker loaded | frequency=%s rows=%d columns=%d original_t0_start=%s original_t0_end=%s review_type_counts=%s",
            self.frequency,
            len(master_df),
            len(master_df.columns),
            original_t0.min().strftime("%Y-%m-%d") if not original_t0.empty else None,
            original_t0.max().strftime("%Y-%m-%d") if not original_t0.empty else None,
            master_df[mh.ReviewType].value_counts(dropna=False).to_dict(),
        )
        self.master_df = master_df



    def calculate_completion_distribution(self, cutoff_days=180):
        sample_st = pd.to_datetime("2026-01-01")
        sample_ed = (self.current_date - pd.Timedelta(
            days=150 + self.current_date.day)).to_period(freq=self.frequency).end_time
        master_df = self.master_df
        period_days = self.frequency_days[self.frequency]

        sample_df = master_df[master_df[mh.OriginalT0].apply(lambda x: sample_st <= x <= sample_ed)]
        sample_df = sample_df.sort_values(by=mh.OriginalT0)
        sample_df[mh.ReviewType] = sample_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in x else "Trigger")
        sample_df['Status'] = sample_df[mh.TaskStatus].map(self.wbh_status)
        sample_df['Case Start'] = master_df[mh.OriginalT0]
        sample_df['Case End'] = sample_df.apply(lambda r: r[mh.ApprovalCancelDate] if r['Status'] == 'Completed' else
                                                r[mh.OriginalT0], axis=1)
        sample_df['Case End'] = sample_df['Case End'].mask(sample_df['Case End'].isna(), master_df[mh.OriginalT0])
        sample_df["N Period"] = (sample_df['Case End'] - sample_df['Case Start']).apply(
            lambda x: x.days // self.frequency_days[self.frequency]
        )
        sample_df = sample_df[sample_df["N Period"].apply(lambda x: 0 <= x <= cutoff_days / period_days - 1)]
        sample_df = sample_df[sample_df["Status"] != "WIP"]

        self.completion_distribution = sample_df.groupby(['Status', "N Period"])[mh.CIN].count() / len(sample_df)
        logger.info(
            "Completion distribution calculated from tracker | frequency=%s sample_rows=%d distribution_points=%d",
            self.frequency,
            len(sample_df),
            len(self.completion_distribution),
        )

    def read_input_completion_percentage(self):
        config = self.input_completion_percentage_config
        completion_distributions = {}
        source_rows = 0
        loaded_case_types = []
        logger.info(
            "Reading completion percentage input | frequency=%s source_frequency=%s file=%s",
            self.frequency,
            config["frequency"],
            self.resolve_input_file_path(config["file_path"]),
        )

        for case_type, input_df in self.read_case_type_input_sheets(config):
            input_df = input_df[[config["period_column"], config["percentage_column"]]].dropna()
            input_df[config["period_column"]] = pd.to_numeric(input_df[config["period_column"]], errors="coerce")
            input_df[config["percentage_column"]] = pd.to_numeric(input_df[config["percentage_column"]], errors="coerce")
            input_df = input_df.dropna()
            source_rows += len(input_df)
            loaded_case_types.append(case_type)

            completion_distribution = input_df.groupby(config["period_column"])[config["percentage_column"]].sum()
            completion_distribution.index = completion_distribution.index.astype(int)
            if len(completion_distribution) > 0 and completion_distribution.max() > 1:
                completion_distribution = completion_distribution / 100

            completion_distribution = self.convert_completion_distribution_frequency(
                completion_distribution,
                input_frequency=config["frequency"]
            )

            for n_period, percentage in completion_distribution.items():
                completion_distributions[(case_type, n_period)] = percentage

        self.completion_distribution = pd.Series(completion_distributions).sort_index()
        if not self.completion_distribution.empty:
            self.completion_distribution.index = pd.MultiIndex.from_tuples(
                self.completion_distribution.index,
                names=["Case Type", "N Period"]
            )
        if self.completion_distribution.empty:
            logger.warning(
                "Completion percentage input produced no distribution | frequency=%s source_rows=%d",
                self.frequency,
                source_rows,
            )
        else:
            probability_totals = self.completion_distribution.groupby(level="Case Type").sum().round(4).to_dict()
            logger.info(
                "Completion percentage input loaded | frequency=%s source_rows=%d distribution_points=%d case_types=%s probability_totals=%s",
                self.frequency,
                source_rows,
                len(self.completion_distribution),
                sorted(set(loaded_case_types)),
                probability_totals,
            )
        return self.completion_distribution

    def convert_completion_distribution_frequency(self, completion_distribution: pd.Series, input_frequency: str):
        input_frequency_days = self.frequency_days[input_frequency]
        output_frequency_days = self.frequency_days[self.frequency]

        completion_distribution = completion_distribution.sort_index()
        max_input_period = int(completion_distribution.index.max())
        max_output_period = math.ceil((max_input_period + 1) * input_frequency_days / output_frequency_days) - 1
        output_distribution = pd.Series(0.0, index=range(max_output_period + 1))

        for input_period, percentage in completion_distribution.items():
            input_start_day = int(input_period) * input_frequency_days
            input_end_day = (int(input_period) + 1) * input_frequency_days

            first_output_period = input_start_day // output_frequency_days
            last_output_period = (input_end_day - 1) // output_frequency_days

            for output_period in range(first_output_period, last_output_period + 1):
                output_start_day = output_period * output_frequency_days
                output_end_day = (output_period + 1) * output_frequency_days
                overlap_days = min(input_end_day, output_end_day) - max(input_start_day, output_start_day)
                output_distribution[output_period] += percentage * overlap_days / input_frequency_days

        output_distribution = output_distribution[output_distribution > 0]
        output_distribution.index.name = "N Period"
        return output_distribution

    def read_input_bow_volume(self):
        config = self.input_bow_volume_config
        input_bow_volume = {}
        output_bow_volume = {}
        source_rows = 0
        loaded_case_types = []
        logger.info(
            "Reading BoW input | frequency=%s source_frequency=%s file=%s",
            self.frequency,
            config["frequency"],
            self.resolve_input_file_path(config["file_path"]),
        )

        for case_type, input_df in self.read_case_type_input_sheets(config):
            input_df = input_df[[config["period_column"], config["volume_column"]]].dropna()
            input_df[config["period_column"]] = pd.to_datetime(input_df[config["period_column"]], errors="coerce")
            input_df[config["volume_column"]] = pd.to_numeric(input_df[config["volume_column"]], errors="coerce")
            input_df = input_df.dropna()
            source_rows += len(input_df)
            loaded_case_types.append(case_type)

            bow_volume = input_df.groupby(config["period_column"])[config["volume_column"]].sum()
            bow_volume.index = pd.to_datetime(bow_volume.index).to_period(config["frequency"])
            bow_volume.index.name = "Input Period"
            bow_volume.name = "Input BoW Volume"

            for input_period, volume in bow_volume.items():
                input_bow_volume[(case_type, input_period)] = volume

            converted_bow_volume = self.convert_bow_volume_frequency(
                bow_volume,
                input_frequency=config["frequency"]
            )

            for period, volume in converted_bow_volume.items():
                output_bow_volume[(case_type, period)] = volume

        self.input_bow_volume = pd.Series(input_bow_volume).sort_index()
        if not self.input_bow_volume.empty:
            self.input_bow_volume.index = pd.MultiIndex.from_tuples(
                self.input_bow_volume.index,
                names=["Case Type", "Input Period"]
            )
        self.input_bow_volume.name = "Input BoW Volume"

        self.bow_volume = pd.Series(output_bow_volume).sort_index()
        if not self.bow_volume.empty:
            self.bow_volume.index = pd.MultiIndex.from_tuples(
                self.bow_volume.index,
                names=["Case Type", "Period"]
            )
        self.bow_volume.name = "BoW Volume"
        logger.info(
            "BoW input loaded | frequency=%s source_rows=%d case_types=%s input_points=%d output_points=%d input_total=%.2f output_total=%.2f",
            self.frequency,
            source_rows,
            sorted(set(loaded_case_types)),
            len(self.input_bow_volume),
            len(self.bow_volume),
            self.numeric_total(self.input_bow_volume),
            self.numeric_total(self.bow_volume),
        )
        return self.bow_volume

    def convert_bow_volume_frequency(self, bow_volume: pd.Series, input_frequency: str):
        output_volume = {}

        for input_period, volume in bow_volume.sort_index().items():
            if not isinstance(input_period, pd.Period):
                input_period = pd.to_datetime(input_period).to_period(input_frequency)
            input_start_date = input_period.start_time.normalize()
            input_end_date = (input_period + 1).start_time.normalize()
            input_days = (input_end_date - input_start_date).days

            first_output_period = input_start_date.to_period(self.frequency)
            last_output_period = (input_end_date - pd.Timedelta(days=1)).to_period(self.frequency)

            for output_period in pd.period_range(first_output_period, last_output_period, freq=self.frequency):
                output_start_date = output_period.start_time.normalize()
                output_end_date = (output_period + 1).start_time.normalize()
                overlap_days = (
                    min(input_end_date, output_end_date) - max(input_start_date, output_start_date)
                ).days

                if overlap_days > 0:
                    output_volume[output_period] = output_volume.get(output_period, 0) + volume * overlap_days / input_days

        output_volume = pd.Series(output_volume).sort_index()
        output_volume.index.name = "Period"
        output_volume.name = "BoW Volume"
        return output_volume

    def resolve_input_file_path(self, file_path):
        if os.path.exists(file_path):
            return file_path

        data_file_path = os.path.join("data", file_path)
        if os.path.exists(data_file_path):
            return data_file_path

        return file_path

    def resolve_input_file_pattern(self, file_path_pattern):
        from glob import glob

        if glob(file_path_pattern):
            return file_path_pattern

        data_file_path_pattern = os.path.join("data", file_path_pattern)
        if glob(data_file_path_pattern):
            return data_file_path_pattern

        return file_path_pattern

    def read_case_type_input_sheets(self, config):
        file_path = self.resolve_input_file_path(config["file_path"])
        excel_file = pd.ExcelFile(file_path)
        sheet_names = config.get("sheet_names")
        if sheet_names is None:
            sheet_names = excel_file.sheet_names
        elif isinstance(sheet_names, str):
            sheet_names = [sheet_names]

        logger.info(
            "Reading input workbook | frequency=%s file=%s sheets=%s",
            self.frequency,
            file_path,
            sheet_names,
        )
        for sheet_name in sheet_names:
            input_df = pd.read_excel(file_path, sheet_name=sheet_name)
            required_columns = {config["period_column"]}
            if "percentage_column" in config:
                required_columns.add(config["percentage_column"])
            if "volume_column" in config:
                required_columns.add(config["volume_column"])

            if not required_columns.issubset(input_df.columns):
                logger.warning(
                    "Skipping input sheet missing columns | frequency=%s file=%s sheet=%s missing_columns=%s",
                    self.frequency,
                    file_path,
                    sheet_name,
                    sorted(required_columns - set(input_df.columns)),
                )
                continue

            input_df = input_df.dropna(how="all")
            if input_df.empty:
                logger.warning(
                    "Skipping empty input sheet | frequency=%s file=%s sheet=%s",
                    self.frequency,
                    file_path,
                    sheet_name,
                )
                continue

            logger.info(
                "Loaded input sheet | frequency=%s file=%s sheet=%s rows=%d columns=%d",
                self.frequency,
                file_path,
                sheet_name,
                len(input_df),
                len(input_df.columns),
            )
            yield sheet_name, input_df

    def get_current_cutoff_date(self):
        return pd.to_datetime(self.current_date).normalize()

    def infer_actual_cutoff_date(self):
        self.actual_cutoff_date = self.get_current_cutoff_date()
        return self.actual_cutoff_date

    def infer_remaining_bow_cutoff_date(self):
        current_date = self.get_current_cutoff_date()
        if mh.OriginalT0 not in self.master_df.columns:
            self.remaining_bow_cutoff_date = current_date
            return self.remaining_bow_cutoff_date

        original_t0 = pd.to_datetime(self.master_df[mh.OriginalT0], errors="coerce").dropna()

        if original_t0.empty:
            self.remaining_bow_cutoff_date = current_date
        else:
            self.remaining_bow_cutoff_date = max(original_t0.max().normalize(), current_date)

        return self.remaining_bow_cutoff_date

    def calculate_received_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_remaining_bow_cutoff_date()

        if mh.OriginalT0 not in self.master_df.columns or mh.ReviewType not in self.master_df.columns:
            self.received_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
            self.received_volume.name = "Received Volume"
            return self.received_volume

        received_df = self.master_df[[mh.OriginalT0, mh.ReviewType]].copy()
        received_df[mh.ReviewType] = received_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
        received_df[mh.OriginalT0] = pd.to_datetime(
            received_df[mh.OriginalT0],
            errors="coerce"
        ).dt.normalize()
        received_df = received_df.dropna(subset=[mh.OriginalT0])
        received_df = received_df[received_df[mh.OriginalT0] <= cutoff_date]
        received_df["Period"] = received_df[mh.OriginalT0].dt.to_period(self.frequency)

        if received_df.empty:
            self.received_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        else:
            self.received_volume = received_df.groupby([mh.ReviewType, "Period"]).size()
            self.received_volume.index = self.received_volume.index.set_names(["Case Type", "Period"])
        self.received_volume.name = "Received Volume"
        self.log_series_summary("Received volume", self.received_volume)
        return self.received_volume

    def calculate_actual_start_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if mh.OriginalT0 not in self.master_df.columns or mh.ReviewType not in self.master_df.columns or mh.TaskStatus not in self.master_df.columns:
            self.actual_start_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
            self.actual_start_volume.name = "Actual Start Volume"
            return self.actual_start_volume

        actual_df = self.master_df[[mh.OriginalT0, mh.ReviewType, mh.TaskStatus]].copy()
        actual_df[mh.OriginalT0] = pd.to_datetime(actual_df[mh.OriginalT0], errors="coerce").dt.normalize()
        actual_df[mh.ReviewType] = actual_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
        actual_df["Status"] = actual_df[mh.TaskStatus].map(self.wbh_status)
        actual_df = actual_df.dropna(subset=[mh.OriginalT0, "Status"])
        actual_df = actual_df[actual_df[mh.OriginalT0] <= cutoff_date]
        actual_df = actual_df[actual_df["Status"].isin(self.actual_start_status)]
        actual_df["Period"] = actual_df[mh.OriginalT0].dt.to_period(self.frequency)

        if actual_df.empty:
            self.actual_start_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        else:
            self.actual_start_volume = actual_df.groupby([mh.ReviewType, "Period"]).size()
            self.actual_start_volume.index = self.actual_start_volume.index.set_names(["Case Type", "Period"])

        self.actual_start_volume.name = "Actual Start Volume"
        self.log_series_summary("Actual start volume", self.actual_start_volume)
        return self.actual_start_volume

    def calculate_open_received_start_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_remaining_bow_cutoff_date()

        if mh.OriginalT0 not in self.master_df.columns or mh.ReviewType not in self.master_df.columns or mh.TaskStatus not in self.master_df.columns:
            self.open_received_start_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Start Period"])
            )
            self.open_received_start_volume.name = "WIP Received Start Volume"
            return self.open_received_start_volume

        start_df = self.master_df[[mh.OriginalT0, mh.ReviewType, mh.TaskStatus]].copy()
        start_df[mh.OriginalT0] = pd.to_datetime(start_df[mh.OriginalT0], errors="coerce").dt.normalize()
        start_df[mh.ReviewType] = start_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
        start_df["Status"] = start_df[mh.TaskStatus].map(self.wbh_status)
        start_df = start_df.dropna(subset=[mh.OriginalT0, "Status"])
        start_df = start_df[start_df[mh.OriginalT0] <= cutoff_date]
        start_df = start_df[start_df["Status"].isin(self.open_start_status)]
        start_df["Start Period"] = start_df[mh.OriginalT0].dt.to_period(self.frequency)

        if start_df.empty:
            self.open_received_start_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Start Period"])
            )
        else:
            self.open_received_start_volume = start_df.groupby([mh.ReviewType, "Start Period"]).size().astype(float)
            self.open_received_start_volume.index = self.open_received_start_volume.index.set_names(["Case Type", "Start Period"])

        self.open_received_start_volume.name = "WIP Received Start Volume"
        self.log_series_summary("WIP received start volume", self.open_received_start_volume)
        return self.open_received_start_volume

    def calculate_pending_qc_ba_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_remaining_bow_cutoff_date()

        if mh.OriginalT0 not in self.master_df.columns or mh.ReviewType not in self.master_df.columns or mh.TaskStatus not in self.master_df.columns:
            self.pending_qc_ba_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Start Period"])
            )
            self.pending_qc_ba_volume.name = "Pending QC/BA Volume"
            return self.pending_qc_ba_volume

        pending_df = self.master_df[[mh.OriginalT0, mh.ReviewType, mh.TaskStatus]].copy()
        pending_df[mh.OriginalT0] = pd.to_datetime(pending_df[mh.OriginalT0], errors="coerce").dt.normalize()
        pending_df[mh.ReviewType] = pending_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
        pending_df["Status"] = pending_df[mh.TaskStatus].map(self.wbh_status)
        pending_df = pending_df.dropna(subset=[mh.OriginalT0, "Status"])
        pending_df = pending_df[pending_df[mh.OriginalT0] <= cutoff_date]
        pending_df = pending_df[pending_df["Status"].isin(self.pending_qc_ba_status)]
        pending_df["Start Period"] = pending_df[mh.OriginalT0].dt.to_period(self.frequency)

        if pending_df.empty:
            self.pending_qc_ba_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Start Period"])
            )
        else:
            self.pending_qc_ba_volume = pending_df.groupby([mh.ReviewType, "Start Period"]).size().astype(float)
            self.pending_qc_ba_volume.index = self.pending_qc_ba_volume.index.set_names(["Case Type", "Start Period"])

        self.pending_qc_ba_volume.name = "Pending QC/BA Volume"
        self.log_series_summary("Pending QC/BA volume", self.pending_qc_ba_volume)
        return self.pending_qc_ba_volume

    def calculate_actual_completion_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if mh.ApprovalCancelDate not in self.master_df.columns or mh.ReviewType not in self.master_df.columns or mh.TaskStatus not in self.master_df.columns:
            self.actual_completion_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
            self.actual_completion_volume.name = "Actual Completion Volume"
            return self.actual_completion_volume

        completion_df = self.master_df[[mh.ApprovalCancelDate, mh.ReviewType, mh.TaskStatus]].copy()
        completion_df[mh.ApprovalCancelDate] = pd.to_datetime(
            completion_df[mh.ApprovalCancelDate],
            errors="coerce"
        ).dt.normalize()
        completion_df[mh.ReviewType] = completion_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
        completion_df["Status"] = completion_df[mh.TaskStatus].map(self.wbh_status)
        completion_df = completion_df.dropna(subset=[mh.ApprovalCancelDate, "Status"])
        completion_df = completion_df[completion_df[mh.ApprovalCancelDate] <= cutoff_date]
        completion_df = completion_df[completion_df["Status"].isin(self.actual_completion_status)]
        completion_df["Period"] = completion_df[mh.ApprovalCancelDate].dt.to_period(self.frequency)

        if completion_df.empty:
            self.actual_completion_volume = pd.Series(
                dtype=int,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        else:
            self.actual_completion_volume = completion_df.groupby([mh.ReviewType, "Period"]).size()
            self.actual_completion_volume.index = self.actual_completion_volume.index.set_names(["Case Type", "Period"])

        self.actual_completion_volume.name = "Actual Completion Volume"
        self.log_series_summary("Actual completion volume", self.actual_completion_volume)
        return self.actual_completion_volume

    def get_case_completion_distribution(self, case_type):
        if self.completion_distribution.empty:
            return pd.Series(dtype=float)

        completion_distribution = self.completion_distribution
        if isinstance(completion_distribution.index, pd.MultiIndex):
            if "Case Type" in completion_distribution.index.names:
                if case_type not in completion_distribution.index.get_level_values("Case Type"):
                    return pd.Series(dtype=float)
                completion_distribution = completion_distribution.xs(case_type, level="Case Type")
            elif "Status" in completion_distribution.index.names:
                if "Completed" not in completion_distribution.index.get_level_values("Status"):
                    return pd.Series(dtype=float)
                completion_distribution = completion_distribution.xs("Completed", level="Status")

        completion_distribution = pd.to_numeric(completion_distribution, errors="coerce").dropna().sort_index()
        completion_distribution.index = completion_distribution.index.astype(int)
        return completion_distribution

    def get_wbh_letter_days(self, case_type):
        normalized_case_type = "PR" if "PR" in str(case_type).upper() else "Trigger"
        return self.wbh_letter_days.get(normalized_case_type)

    def get_wbh_call_days(self, case_type):
        normalized_case_type = "PR" if "PR" in str(case_type).upper() else "Trigger"
        return self.wbh_call_days.get(normalized_case_type)

    def calculate_uncompleted_probability(self, case_type, elapsed_periods):
        completion_distribution = self.get_case_completion_distribution(case_type)
        if completion_distribution.empty:
            return None

        completed_probability = completion_distribution[completion_distribution.index < elapsed_periods].sum()
        return max(1 - completed_probability, 0)

    def add_wbh_letter_volume(self, wbh_letter_volume, case_type, start_period, start_volume, source,
                              condition_on_cutoff=False, cutoff_date=None):
        if start_volume <= 0:
            return

        letter_days = self.get_wbh_letter_days(case_type)
        if letter_days is None:
            return

        letter_days_to_period = letter_days // self.frequency_days[self.frequency]
        letter_period = start_period + letter_days_to_period
        letter_elapsed_periods = max(letter_period.ordinal - start_period.ordinal, 0)
        letter_probability = self.calculate_uncompleted_probability(case_type, letter_elapsed_periods)
        if letter_probability is None or letter_probability <= 0:
            return

        output_period = letter_period
        if condition_on_cutoff and cutoff_date is not None:
            cutoff_period = pd.to_datetime(cutoff_date).to_period(self.frequency)
            cutoff_elapsed_periods = max(cutoff_period.ordinal - start_period.ordinal, 0)

            if letter_elapsed_periods <= cutoff_elapsed_periods:
                letter_probability = 1
            else:
                cutoff_survival_probability = self.calculate_uncompleted_probability(
                    case_type,
                    cutoff_elapsed_periods
                )
                if cutoff_survival_probability is None or cutoff_survival_probability <= 0:
                    return
                letter_probability = letter_probability / cutoff_survival_probability

        wbh_letter_volume[(case_type, output_period, source)] = (
            wbh_letter_volume.get((case_type, output_period, source), 0)
            + start_volume * letter_probability
        )

    def calculate_wbh_letter_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if self.completion_distribution.empty:
            self.read_input_completion_percentage()

        if self.remaining_bow_volume.empty:
            self.calculate_remaining_bow_volume()

        open_received_cutoff_date = self.infer_remaining_bow_cutoff_date()
        open_received_start_volume = self.calculate_open_received_start_volume(cutoff_date=open_received_cutoff_date)
        wbh_letter_volume = {}

        for (case_type, start_period), start_volume in open_received_start_volume.items():
            self.add_wbh_letter_volume(
                wbh_letter_volume,
                case_type,
                start_period,
                start_volume,
                source="WIP Received",
                condition_on_cutoff=True,
                cutoff_date=cutoff_date
            )

        if not self.remaining_bow_volume.empty:
            remaining_start_volume = self.remaining_bow_volume["Remaining BoW Volume"]
            for (case_type, start_period), start_volume in remaining_start_volume.items():
                self.add_wbh_letter_volume(
                    wbh_letter_volume,
                    case_type,
                    start_period,
                    start_volume,
                    source="Remaining BoW",
                    condition_on_cutoff=False,
                    cutoff_date=cutoff_date
                )

        if wbh_letter_volume:
            wbh_letter_volume = pd.Series(wbh_letter_volume, dtype=float).sort_index()
            wbh_letter_volume.index = pd.MultiIndex.from_tuples(
                wbh_letter_volume.index,
                names=["Case Type", "Period", "Source"]
            )
        else:
            wbh_letter_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], [], []], names=["Case Type", "Period", "Source"])
            )

        source_volume = wbh_letter_volume.unstack("Source", fill_value=0)
        periods = source_volume.index.sort_values()
        wbh_letter_df = pd.DataFrame(index=periods)
        wbh_letter_df.index = wbh_letter_df.index.set_names(["Case Type", "Period"])
        wbh_letter_df["WIP Received WBH Letter Volume"] = source_volume.get("WIP Received", 0)
        wbh_letter_df["Remaining BoW WBH Letter Volume"] = source_volume.get("Remaining BoW", 0)
        wbh_letter_df["Forecast WBH Letter Volume"] = (
            wbh_letter_df["WIP Received WBH Letter Volume"]
            + wbh_letter_df["Remaining BoW WBH Letter Volume"]
        )

        self.forecast_wbh_letter_volume = wbh_letter_df["Forecast WBH Letter Volume"]
        self.forecast_wbh_letter_volume.name = "Forecast WBH Letter Volume"
        self.wbh_letter_volume = wbh_letter_df
        logger.info(
            "WBH letter summary | frequency=%s rows=%d total=%.2f",
            self.frequency,
            len(self.wbh_letter_volume),
            float(pd.to_numeric(self.forecast_wbh_letter_volume, errors="coerce").fillna(0).sum()),
        )
        return self.wbh_letter_volume

    def add_wbh_call_volume(self, wbh_call_volume, case_type, start_period, start_volume, source,
                            condition_on_cutoff=False, cutoff_date=None):
        if start_volume <= 0:
            return

        call_days = self.get_wbh_call_days(case_type)
        if call_days is None:
            return

        call_days_to_period = call_days // self.frequency_days[self.frequency]
        call_period = start_period + call_days_to_period
        call_elapsed_periods = max(call_period.ordinal - start_period.ordinal, 0)
        call_probability = self.calculate_uncompleted_probability(case_type, call_elapsed_periods)
        if call_probability is None or call_probability <= 0:
            return

        output_period = call_period
        if condition_on_cutoff and cutoff_date is not None:
            cutoff_period = pd.to_datetime(cutoff_date).to_period(self.frequency)
            cutoff_elapsed_periods = max(cutoff_period.ordinal - start_period.ordinal, 0)

            if call_elapsed_periods <= cutoff_elapsed_periods:
                call_probability = 1
            else:
                cutoff_survival_probability = self.calculate_uncompleted_probability(
                    case_type,
                    cutoff_elapsed_periods
                )
                if cutoff_survival_probability is None or cutoff_survival_probability <= 0:
                    return
                call_probability = call_probability / cutoff_survival_probability

        wbh_call_volume[(case_type, output_period, source)] = (
            wbh_call_volume.get((case_type, output_period, source), 0)
            + start_volume * call_probability
        )

    def calculate_wbh_call_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if self.completion_distribution.empty:
            self.read_input_completion_percentage()

        if self.remaining_bow_volume.empty:
            self.calculate_remaining_bow_volume()

        open_received_cutoff_date = self.infer_remaining_bow_cutoff_date()
        open_received_start_volume = self.calculate_open_received_start_volume(cutoff_date=open_received_cutoff_date)
        wbh_call_volume = {}

        for (case_type, start_period), start_volume in open_received_start_volume.items():
            self.add_wbh_call_volume(
                wbh_call_volume,
                case_type,
                start_period,
                start_volume,
                source="WIP Received",
                condition_on_cutoff=True,
                cutoff_date=cutoff_date
            )

        if not self.remaining_bow_volume.empty:
            remaining_start_volume = self.remaining_bow_volume["Remaining BoW Volume"]
            for (case_type, start_period), start_volume in remaining_start_volume.items():
                self.add_wbh_call_volume(
                    wbh_call_volume,
                    case_type,
                    start_period,
                    start_volume,
                    source="Remaining BoW",
                    condition_on_cutoff=False,
                    cutoff_date=cutoff_date
                )

        if wbh_call_volume:
            wbh_call_volume = pd.Series(wbh_call_volume, dtype=float).sort_index()
            wbh_call_volume.index = pd.MultiIndex.from_tuples(
                wbh_call_volume.index,
                names=["Case Type", "Period", "Source"]
            )
        else:
            wbh_call_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], [], []], names=["Case Type", "Period", "Source"])
            )

        source_volume = wbh_call_volume.unstack("Source", fill_value=0)
        periods = source_volume.index.sort_values()
        wbh_call_df = pd.DataFrame(index=periods)
        wbh_call_df.index = wbh_call_df.index.set_names(["Case Type", "Period"])
        wbh_call_df["WIP Received WBH Call Volume"] = source_volume.get("WIP Received", 0)
        wbh_call_df["Remaining BoW WBH Call Volume"] = source_volume.get("Remaining BoW", 0)
        wbh_call_df["Forecast WBH Call Volume"] = (
            wbh_call_df["WIP Received WBH Call Volume"]
            + wbh_call_df["Remaining BoW WBH Call Volume"]
        )

        self.forecast_wbh_call_volume = wbh_call_df["Forecast WBH Call Volume"]
        self.forecast_wbh_call_volume.name = "Forecast WBH Call Volume"
        self.wbh_call_volume = wbh_call_df
        logger.info(
            "WBH call summary | frequency=%s rows=%d total=%.2f",
            self.frequency,
            len(self.wbh_call_volume),
            float(pd.to_numeric(self.forecast_wbh_call_volume, errors="coerce").fillna(0).sum()),
        )
        return self.wbh_call_volume

    def add_forecast_completion_volume(self, forecast_completion_volume, case_type, start_period, start_volume,
                                       condition_on_cutoff=False, cutoff_date=None):
        completion_distribution = self.get_case_completion_distribution(case_type)
        if start_volume <= 0 or completion_distribution.empty:
            return

        elapsed_periods = 0
        survival_probability = 1
        if condition_on_cutoff and cutoff_date is not None:
            cutoff_period = pd.to_datetime(cutoff_date).to_period(self.frequency)
            elapsed_periods = max(cutoff_period.ordinal - start_period.ordinal, 0)
            survival_probability = 1 - completion_distribution[completion_distribution.index < elapsed_periods].sum()
            if survival_probability <= 0:
                return

        for n_period, percentage in completion_distribution.items():

            n_period = int(n_period)
            if condition_on_cutoff and n_period < elapsed_periods:
                continue

            completion_period = start_period + n_period
            adjusted_percentage = percentage / survival_probability if condition_on_cutoff else percentage
            forecast_completion_volume[(case_type, completion_period)] = (
                forecast_completion_volume.get((case_type, completion_period), 0)
                + start_volume * adjusted_percentage
            )

    def calculate_completion_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if self.completion_distribution.empty:
            self.read_input_completion_percentage()

        if self.remaining_bow_volume.empty:
            self.calculate_remaining_bow_volume()

        open_received_cutoff_date = self.infer_remaining_bow_cutoff_date()
        open_received_start_volume = self.calculate_open_received_start_volume(cutoff_date=open_received_cutoff_date)
        forecast_completion_volume = {}

        for (case_type, start_period), start_volume in open_received_start_volume.items():
            self.add_forecast_completion_volume(
                forecast_completion_volume,
                case_type,
                start_period,
                start_volume,
                condition_on_cutoff=True,
                cutoff_date=cutoff_date
            )

        if not self.remaining_bow_volume.empty:
            remaining_start_volume = self.remaining_bow_volume["Remaining BoW Volume"]
            for (case_type, start_period), start_volume in remaining_start_volume.items():
                self.add_forecast_completion_volume(
                    forecast_completion_volume,
                    case_type,
                    start_period,
                    start_volume,
                    condition_on_cutoff=False,
                    cutoff_date=cutoff_date
                )

        if forecast_completion_volume:
            self.forecast_completion_volume = pd.Series(forecast_completion_volume, dtype=float).sort_index()
            self.forecast_completion_volume.index = pd.MultiIndex.from_tuples(
                self.forecast_completion_volume.index,
                names=["Case Type", "Period"]
            )
        else:
            self.forecast_completion_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        self.forecast_completion_volume.name = "Forecast Completion Volume"

        self.calculate_actual_completion_volume(cutoff_date=cutoff_date)

        periods = self.forecast_completion_volume.index.union(self.actual_completion_volume.index).sort_values()
        completion_df = pd.DataFrame(index=periods)
        completion_df.index = completion_df.index.set_names(["Case Type", "Period"])
        completion_df["Forecast Completion Volume"] = self.forecast_completion_volume.reindex(periods, fill_value=0)
        completion_df["Actual Completion Volume"] = self.actual_completion_volume.reindex(periods, fill_value=0)

        self.completion_volume = completion_df
        logger.info(
            "Completion volume summary | frequency=%s rows=%d forecast_total=%.2f actual_total=%.2f",
            self.frequency,
            len(self.completion_volume),
            float(pd.to_numeric(self.completion_volume["Forecast Completion Volume"], errors="coerce").fillna(0).sum()),
            float(pd.to_numeric(self.completion_volume["Actual Completion Volume"], errors="coerce").fillna(0).sum()),
        )
        return self.completion_volume

    def calculate_remaining_bow_volume(self, cutoff_date=None):
        if self.input_bow_volume.empty:
            self.read_input_bow_volume()

        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()
        actual_cutoff_date = pd.to_datetime(cutoff_date).normalize()
        bow_cutoff_date = self.infer_remaining_bow_cutoff_date()
        self.calculate_received_volume(cutoff_date=bow_cutoff_date)

        input_frequency = self.input_bow_volume_config["frequency"]
        received_by_input_period = pd.Series(dtype=int)
        if (
            self.frequency in {"W", "D"}
            and mh.OriginalT0 in self.master_df.columns
            and mh.ReviewType in self.master_df.columns
        ):
            received_df = self.master_df[[mh.OriginalT0, mh.ReviewType]].copy()
            received_df[mh.ReviewType] = received_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
            received_df[mh.OriginalT0] = pd.to_datetime(received_df[mh.OriginalT0], errors="coerce").dt.normalize()
            received_df = received_df.dropna(subset=[mh.OriginalT0])
            received_df = received_df[received_df[mh.OriginalT0] <= actual_cutoff_date]
            received_df["Input Period"] = received_df[mh.OriginalT0].dt.to_period(input_frequency)
            received_by_input_period = received_df.groupby([mh.ReviewType, "Input Period"]).size()
            received_by_input_period.index = received_by_input_period.index.set_names(["Case Type", "Input Period"])

        output_plan_volume = {}

        for (case_type, input_period), planned_volume in self.input_bow_volume.sort_index().items():
            input_start_date = input_period.start_time.normalize()
            input_end_date = (input_period + 1).start_time.normalize()

            if input_end_date <= actual_cutoff_date:
                continue

            allocation_start_date = max(input_start_date, actual_cutoff_date)
            allocation_volume = planned_volume

            if self.frequency in {"W", "D"}:
                received_volume = received_by_input_period.get((case_type, input_period), 0)
                allocation_volume = max(0, planned_volume - received_volume)

            if allocation_volume == 0 or allocation_start_date >= input_end_date:
                continue

            allocation_days = (input_end_date - allocation_start_date).days
            first_output_period = allocation_start_date.to_period(self.frequency)
            last_output_period = (input_end_date - pd.Timedelta(days=1)).to_period(self.frequency)

            for output_period in pd.period_range(first_output_period, last_output_period, freq=self.frequency):
                output_start_date = output_period.start_time.normalize()
                output_end_date = (output_period + 1).start_time.normalize()
                overlap_days = (
                    min(input_end_date, output_end_date) - max(allocation_start_date, output_start_date)
                ).days

                if overlap_days > 0:
                    output_plan_volume[(case_type, output_period)] = (
                        output_plan_volume.get((case_type, output_period), 0)
                        + allocation_volume * overlap_days / allocation_days
                    )

        if output_plan_volume:
            output_plan_volume = pd.Series(output_plan_volume, dtype=float).sort_index()
            output_plan_volume.index = pd.MultiIndex.from_tuples(
                output_plan_volume.index,
                names=["Case Type", "Period"]
            )
        else:
            output_plan_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        output_plan_volume.name = "Input BoW Volume"

        current_output_period = actual_cutoff_date.to_period(self.frequency)
        current_and_future_received = self.received_volume[
            self.received_volume.index.get_level_values("Period") >= current_output_period
        ]
        periods = output_plan_volume.index.union(current_and_future_received.index).sort_values()
        remaining_df = pd.DataFrame(index=periods)
        remaining_df.index = remaining_df.index.set_names(["Case Type", "Period"])
        remaining_df["Input BoW Volume"] = output_plan_volume.reindex(periods, fill_value=0)
        remaining_df["Received Volume"] = current_and_future_received.reindex(periods, fill_value=0)
        remaining_df["Remaining BoW Volume"] = (
            remaining_df["Input BoW Volume"] - remaining_df["Received Volume"]
        ).clip(lower=0)
        remaining_df["Over Received Volume"] = (
            remaining_df["Received Volume"] - remaining_df["Input BoW Volume"]
        ).clip(lower=0)

        self.remaining_bow_volume = remaining_df
        logger.info(
            "Remaining BoW summary | frequency=%s rows=%d input_total=%.2f received_total=%.2f remaining_total=%.2f over_received_total=%.2f",
            self.frequency,
            len(self.remaining_bow_volume),
            float(pd.to_numeric(self.remaining_bow_volume["Input BoW Volume"], errors="coerce").fillna(0).sum()),
            float(pd.to_numeric(self.remaining_bow_volume["Received Volume"], errors="coerce").fillna(0).sum()),
            float(pd.to_numeric(self.remaining_bow_volume["Remaining BoW Volume"], errors="coerce").fillna(0).sum()),
            float(pd.to_numeric(self.remaining_bow_volume["Over Received Volume"], errors="coerce").fillna(0).sum()),
        )
        return self.remaining_bow_volume["Remaining BoW Volume"]


    def calculate_workload(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        received_start_volume = self.calculate_received_volume(
            cutoff_date=self.infer_remaining_bow_cutoff_date()
        )
        open_received_start_volume = self.calculate_open_received_start_volume(
            cutoff_date=self.infer_remaining_bow_cutoff_date()
        )
        pending_qc_ba_volume = self.calculate_pending_qc_ba_volume(
            cutoff_date=self.infer_remaining_bow_cutoff_date()
        )
        remaining_start_volume = self.calculate_remaining_bow_volume(cutoff_date=cutoff_date)
        actual_start_volume = self.calculate_actual_start_volume(cutoff_date=cutoff_date)
        completion_volume = self.calculate_completion_volume(cutoff_date=cutoff_date)
        wbh_letter_volume = self.calculate_wbh_letter_volume(cutoff_date=cutoff_date)
        wbh_call_volume = self.calculate_wbh_call_volume(cutoff_date=cutoff_date)

        empty_index = pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])

        def normalize_volume_series(volume_series, name):
            if volume_series is None or volume_series.empty:
                return pd.Series(dtype=float, index=empty_index, name=name)

            volume_series = volume_series.copy()
            volume_series.index = volume_series.index.set_names(["Case Type", "Period"])
            volume_series = pd.to_numeric(volume_series, errors="coerce").fillna(0).astype(float)
            volume_series.name = name
            return volume_series

        workload_series = {
            "Planned Start Volume": normalize_volume_series(self.bow_volume, "Planned Start Volume"),
            "Received Start Volume": normalize_volume_series(received_start_volume, "Received Start Volume"),
            "WIP Received Start Volume": normalize_volume_series(
                open_received_start_volume,
                "WIP Received Start Volume"
            ),
            "Pending QC/BA Volume": normalize_volume_series(
                pending_qc_ba_volume,
                "Pending QC/BA Volume"
            ),
            "Remaining Planned Start Volume": normalize_volume_series(
                remaining_start_volume,
                "Remaining Planned Start Volume"
            ),
            "Actual Start Volume": normalize_volume_series(actual_start_volume, "Actual Start Volume"),
            "Actual Completion Volume": normalize_volume_series(
                completion_volume.get("Actual Completion Volume") if not completion_volume.empty else None,
                "Actual Completion Volume"
            ),
            "Forecast Completion Volume": normalize_volume_series(
                completion_volume.get("Forecast Completion Volume") if not completion_volume.empty else None,
                "Forecast Completion Volume"
            ),
            "Forecast WBH Letter Volume": normalize_volume_series(
                wbh_letter_volume.get("Forecast WBH Letter Volume") if not wbh_letter_volume.empty else None,
                "Forecast WBH Letter Volume"
            ),
            "Forecast WBH Call Volume": normalize_volume_series(
                wbh_call_volume.get("Forecast WBH Call Volume") if not wbh_call_volume.empty else None,
                "Forecast WBH Call Volume"
            ),
        }

        periods = empty_index
        for volume_series in workload_series.values():
            periods = periods.union(volume_series.index)
        periods = periods.sort_values()

        workload_df = pd.DataFrame(index=periods)
        workload_df.index = workload_df.index.set_names(["Case Type", "Period"])
        for column, volume_series in workload_series.items():
            workload_df[column] = volume_series.reindex(periods, fill_value=0)

        workload_df["Completion Volume"] = (
            workload_df["Actual Completion Volume"] + workload_df["Forecast Completion Volume"]
        )
        workload_df["Init Volume"] = (
            workload_df["Received Start Volume"] + workload_df["Remaining Planned Start Volume"]
        )
        workload_df["Demand FTE"] = np.nan
        if self.working_hour:
            workload_df["Demand FTE"] = (
                workload_df["Completion Volume"] * self.completion_upt
                + workload_df["Init Volume"] * self.init_upt
            ) / self.working_hour

        self.workload_volume = workload_df
        self.log_workload_summary()
        return self.workload_volume

    def run(self):
        start_time = time.perf_counter()
        logger.info(
            "Data processor run start | frequency=%s current_date=%s",
            self.frequency,
            pd.to_datetime(self.current_date).strftime("%Y-%m-%d"),
        )
        self.read_data()
        if self.read_completion_percentage_from_input:
            self.read_input_completion_percentage()
        else:
            self.calculate_completion_distribution()
        self.read_input_bow_volume()
        self.calculate_workload()
        logger.info(
            "Data processor run complete | frequency=%s elapsed_sec=%.2f",
            self.frequency,
            time.perf_counter() - start_time,
        )
```
