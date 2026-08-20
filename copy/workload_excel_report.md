# workload_excel_report.py

```python
import os
import logging

import numpy as np
import pandas as pd

try:
    from .config_loader import load_config
except ImportError:
    from config_loader import load_config

logger = logging.getLogger(__name__)


class WorkloadExcelReporter:

    def __init__(self, processor):
        self.processor = processor
        self.config = getattr(processor, "config", None) or load_config()
        self.output_excel_tables = {}
        self.output_excel_path = None
        self.forecast_completion_calculation = pd.DataFrame()
        self.wbh_letter_calculation = pd.DataFrame()
        self.wbh_call_calculation = pd.DataFrame()
        self.demand_fte_calculation = pd.DataFrame()

    def add_period_display_columns(self, output_df, period_columns=None):
        if period_columns is None:
            period_columns = ["Period"]

        output_df = output_df.copy()
        for period_column in period_columns:
            if period_column not in output_df.columns:
                continue

            period_start_column = f"{period_column} Start"
            period_end_column = f"{period_column} End"

            output_df[period_start_column] = output_df[period_column].apply(
                lambda period: period.start_time.normalize() if isinstance(period, pd.Period) else pd.NaT
            )
            output_df[period_end_column] = output_df[period_column].apply(
                lambda period: period.end_time.normalize() if isinstance(period, pd.Period) else pd.NaT
            )
            output_df[period_column] = output_df[period_column].astype(str)

        return output_df

    def output_table_from_series(self, volume_series, value_name, period_columns=None):
        if volume_series is None or volume_series.empty:
            return pd.DataFrame()

        output_df = volume_series.reset_index(name=value_name)
        return self.add_period_display_columns(output_df, period_columns=period_columns)

    def output_table_from_dataframe(self, volume_df, period_columns=None):
        if volume_df is None or volume_df.empty:
            return pd.DataFrame()

        output_df = volume_df.reset_index()
        return self.add_period_display_columns(output_df, period_columns=period_columns)

    def build_start_reconciliation_output_table(self, actual_cutoff_date, bow_cutoff_date):
        output_df = self.output_table_from_dataframe(
            self.processor.remaining_bow_volume,
            period_columns=["Period"]
        )
        if output_df.empty:
            return output_df

        output_df = output_df.rename(
            columns={
                "Input BoW Volume": "Output Plan Volume",
                "Received Volume": "Received Through BoW Cutoff",
            }
        )
        output_df["Actual Cutoff Date"] = pd.to_datetime(actual_cutoff_date).normalize()
        output_df["BoW Cutoff Date"] = pd.to_datetime(bow_cutoff_date).normalize()
        output_df["Reconciliation Difference"] = (
            output_df["Output Plan Volume"]
            - output_df["Received Through BoW Cutoff"]
            - output_df["Remaining BoW Volume"]
            + output_df["Over Received Volume"]
        )
        output_df["Reconciliation Status"] = np.where(
            output_df["Reconciliation Difference"].abs() < 0.000001,
            "OK",
            "CHECK",
        )

        return output_df[
            [
                "Case Type",
                "Period",
                "Period Start",
                "Period End",
                "Actual Cutoff Date",
                "BoW Cutoff Date",
                "Output Plan Volume",
                "Received Through BoW Cutoff",
                "Remaining BoW Volume",
                "Over Received Volume",
                "Reconciliation Difference",
                "Reconciliation Status",
            ]
        ]

    def build_completion_distribution_output_table(self):
        processor = self.processor
        distribution_df = self.output_table_from_series(
            processor.completion_distribution,
            "Completion Probability",
            period_columns=[]
        )
        if distribution_df.empty:
            return distribution_df

        if {"Case Type", "N Period", "Completion Probability"}.issubset(distribution_df.columns):
            distribution_df = distribution_df.sort_values(["Case Type", "N Period"])
            distribution_df["Cumulative Completion Probability"] = (
                distribution_df.groupby("Case Type")["Completion Probability"].cumsum()
            )
            distribution_df["Remaining Uncompleted Probability"] = (
                1 - distribution_df["Cumulative Completion Probability"]
            ).clip(lower=0)

        return distribution_df

    def get_forecast_start_volume_inputs(self):
        processor = self.processor
        if processor.remaining_bow_volume.empty:
            processor.calculate_remaining_bow_volume()

        master_file_cutoff_date = processor.infer_remaining_bow_cutoff_date()
        current_cutoff_date = processor.get_current_cutoff_date()
        wip_received_start_volume = processor.calculate_open_received_start_volume(
            cutoff_date=master_file_cutoff_date
        )
        start_inputs = []

        for (case_type, start_period), start_volume in wip_received_start_volume.items():
            if start_volume <= 0:
                continue

            start_inputs.append(
                {
                    "Case Type": case_type,
                    "Source": "WIP Received",
                    "Start Period": start_period,
                    "Start Volume": float(start_volume),
                    "Conditioned on Master File Cutoff": True,
                    "Cutoff Date": current_cutoff_date,
                }
            )

        if not processor.remaining_bow_volume.empty:
            remaining_start_volume = processor.remaining_bow_volume["Remaining BoW Volume"]
            for (case_type, start_period), start_volume in remaining_start_volume.items():
                if start_volume <= 0:
                    continue

                start_inputs.append(
                    {
                        "Case Type": case_type,
                        "Source": "Remaining Planned",
                        "Start Period": start_period,
                        "Start Volume": float(start_volume),
                        "Conditioned on Master File Cutoff": False,
                        "Cutoff Date": pd.NaT,
                    }
                )

        return start_inputs

    def build_forecast_completion_calculation_table(self):
        processor = self.processor
        records = []

        for start_input in self.get_forecast_start_volume_inputs():
            case_type = start_input["Case Type"]
            start_period = start_input["Start Period"]
            start_volume = start_input["Start Volume"]
            conditioned_on_cutoff = start_input["Conditioned on Master File Cutoff"]
            cutoff_date = start_input["Cutoff Date"]
            completion_distribution = processor.get_case_completion_distribution(case_type)
            if completion_distribution.empty:
                continue

            cutoff_period = pd.NaT
            cutoff_elapsed_periods = 0
            completed_probability_before_cutoff = 0
            survival_probability = 1
            if conditioned_on_cutoff and pd.notna(cutoff_date):
                cutoff_period = pd.to_datetime(cutoff_date).to_period(processor.frequency)
                cutoff_elapsed_periods = max(cutoff_period.ordinal - start_period.ordinal, 0)
                completed_probability_before_cutoff = completion_distribution[
                    completion_distribution.index < cutoff_elapsed_periods
                ].sum()
                survival_probability = 1 - completed_probability_before_cutoff

            for n_period, raw_probability in completion_distribution.items():
                n_period = int(n_period)
                completion_period = start_period + n_period
                included_in_forecast = survival_probability > 0 and (
                    not conditioned_on_cutoff or n_period >= cutoff_elapsed_periods
                )
                applied_probability = (
                    raw_probability / survival_probability
                    if included_in_forecast and conditioned_on_cutoff
                    else raw_probability if included_in_forecast else 0
                )

                records.append(
                    {
                        "Case Type": case_type,
                        "Source": start_input["Source"],
                        "Start Period": start_period,
                        "Start Volume": start_volume,
                        "Conditioned on Master File Cutoff": conditioned_on_cutoff,
                        "Cutoff Date": cutoff_date,
                        "Cutoff Period": cutoff_period,
                        "Cutoff Elapsed Periods": cutoff_elapsed_periods,
                        "Completed Probability Before Cutoff": completed_probability_before_cutoff,
                        "Survival Probability at Cutoff": survival_probability,
                        "N Period": n_period,
                        "Forecast Completion Period": completion_period,
                        "Raw Completion Probability": raw_probability,
                        "Applied Completion Probability": applied_probability,
                        "Forecast Completion Volume": start_volume * applied_probability,
                        "Included In Forecast": included_in_forecast,
                    }
                )

        calculation_df = pd.DataFrame(records)
        calculation_df = self.add_period_display_columns(
            calculation_df,
            period_columns=["Start Period", "Cutoff Period", "Forecast Completion Period"]
        )
        self.forecast_completion_calculation = calculation_df
        setattr(processor, "forecast_completion_calculation", calculation_df)
        return self.forecast_completion_calculation

    def build_wbh_action_calculation_table(self, action_name):
        processor = self.processor
        records = []
        if action_name == "WBH Letter":
            days_getter = processor.get_wbh_letter_days
            volume_column = "Forecast WBH Letter Volume"
        elif action_name == "WBH Call":
            days_getter = processor.get_wbh_call_days
            volume_column = "Forecast WBH Call Volume"
        else:
            raise ValueError(f"Unsupported WBH action: {action_name}")

        for start_input in self.get_forecast_start_volume_inputs():
            case_type = start_input["Case Type"]
            start_period = start_input["Start Period"]
            start_volume = start_input["Start Volume"]
            conditioned_on_cutoff = start_input["Conditioned on Master File Cutoff"]
            cutoff_date = start_input["Cutoff Date"]
            action_days = days_getter(case_type)
            if action_days is None:
                continue

            action_periods = action_days // processor.frequency_days[processor.frequency]
            action_period = start_period + action_periods
            action_elapsed_periods = max(action_period.ordinal - start_period.ordinal, 0)
            uncompleted_probability_at_action = processor.calculate_uncompleted_probability(
                case_type,
                action_elapsed_periods
            )
            if uncompleted_probability_at_action is None:
                continue

            completed_probability_before_action = 1 - uncompleted_probability_at_action
            cutoff_period = pd.NaT
            cutoff_elapsed_periods = 0
            uncompleted_probability_at_cutoff = 1
            applied_probability = uncompleted_probability_at_action
            conditioning_note = "Not conditioned"

            if conditioned_on_cutoff and pd.notna(cutoff_date):
                cutoff_period = pd.to_datetime(cutoff_date).to_period(processor.frequency)
                cutoff_elapsed_periods = max(cutoff_period.ordinal - start_period.ordinal, 0)
                uncompleted_probability_at_cutoff = processor.calculate_uncompleted_probability(
                    case_type,
                    cutoff_elapsed_periods
                )
                if action_elapsed_periods <= cutoff_elapsed_periods:
                    applied_probability = 1
                    conditioning_note = "Action period reached by Master File cutoff"
                elif uncompleted_probability_at_cutoff and uncompleted_probability_at_cutoff > 0:
                    applied_probability = uncompleted_probability_at_action / uncompleted_probability_at_cutoff
                    conditioning_note = "Conditioned on survival to Master File cutoff"
                else:
                    applied_probability = 0
                    conditioning_note = "No survival probability at Master File cutoff"

            records.append(
                {
                    "Case Type": case_type,
                    "Source": start_input["Source"],
                    "Start Period": start_period,
                    "Start Volume": start_volume,
                    "Action": action_name,
                    "Rule Days": action_days,
                    "Rule Periods": action_periods,
                    "Action Period": action_period,
                    "Action Elapsed Periods": action_elapsed_periods,
                    "Completed Probability Before Action": completed_probability_before_action,
                    "Uncompleted Probability at Action": uncompleted_probability_at_action,
                    "Conditioned on Master File Cutoff": conditioned_on_cutoff,
                    "Cutoff Date": cutoff_date,
                    "Cutoff Period": cutoff_period,
                    "Cutoff Elapsed Periods": cutoff_elapsed_periods,
                    "Uncompleted Probability at Cutoff": uncompleted_probability_at_cutoff,
                    "Applied Action Probability": applied_probability,
                    volume_column: start_volume * applied_probability,
                    "Conditioning Note": conditioning_note,
                }
            )

        calculation_df = pd.DataFrame(records)
        calculation_df = self.add_period_display_columns(
            calculation_df,
            period_columns=["Start Period", "Action Period", "Cutoff Period"]
        )

        if action_name == "WBH Letter":
            self.wbh_letter_calculation = calculation_df
            setattr(processor, "wbh_letter_calculation", calculation_df)
            return self.wbh_letter_calculation

        self.wbh_call_calculation = calculation_df
        setattr(processor, "wbh_call_calculation", calculation_df)
        return self.wbh_call_calculation

    def build_demand_fte_calculation_table(self):
        processor = self.processor
        if processor.workload_volume.empty:
            processor.calculate_workload()

        if processor.workload_volume.empty:
            return pd.DataFrame()

        calculation_df = processor.workload_volume.reset_index().copy()
        volume_columns = [
            "Actual Completion Volume",
            "Forecast Completion Volume",
            "Completion Volume",
            "Actual Start Volume",
            "Remaining Planned Start Volume",
            "Init Volume",
        ]
        for volume_column in volume_columns:
            if volume_column not in calculation_df.columns:
                calculation_df[volume_column] = 0
            calculation_df[volume_column] = pd.to_numeric(
                calculation_df[volume_column],
                errors="coerce"
            ).fillna(0)

        calculation_df["Completion UPT"] = processor.completion_upt
        calculation_df["Init UPT"] = processor.init_upt
        calculation_df["Working Hour"] = processor.working_hour
        calculation_df["Completion Hours"] = (
            calculation_df["Completion Volume"] * calculation_df["Completion UPT"]
        )
        calculation_df["Init Hours"] = calculation_df["Init Volume"] * calculation_df["Init UPT"]
        calculation_df["Demand Hours"] = (
            calculation_df["Completion Hours"] + calculation_df["Init Hours"]
        )
        calculation_df["Demand FTE"] = np.nan
        if processor.working_hour:
            calculation_df["Demand FTE"] = calculation_df["Demand Hours"] / processor.working_hour
        calculation_df["Formula"] = (
            "(Completion Volume * Completion UPT + Init Volume * Init UPT) / Working Hour"
        )

        calculation_df = self.add_period_display_columns(calculation_df, period_columns=["Period"])
        calculation_df = calculation_df[
            [
                "Case Type",
                "Period",
                "Period Start",
                "Period End",
                "Actual Completion Volume",
                "Forecast Completion Volume",
                "Completion Volume",
                "Actual Start Volume",
                "Remaining Planned Start Volume",
                "Init Volume",
                "Completion UPT",
                "Init UPT",
                "Working Hour",
                "Completion Hours",
                "Init Hours",
                "Demand Hours",
                "Demand FTE",
                "Formula",
            ]
        ]

        self.demand_fte_calculation = calculation_df
        setattr(processor, "demand_fte_calculation", calculation_df)
        return self.demand_fte_calculation

    def format_rule_days(self, rule_days):
        return ", ".join(
            f"{case_type} T+{days}"
            for case_type, days in rule_days.items()
        )

    def build_output_excel_tables(self, cutoff_date=None, output_path=None, write_excel=True):
        processor = self.processor
        if cutoff_date is None:
            cutoff_date = processor.infer_actual_cutoff_date()
        actual_cutoff_date = pd.to_datetime(cutoff_date).normalize()

        logger.info(
            "Building Excel output tables | frequency=%s cutoff_date=%s write_excel=%s output_path=%s",
            processor.frequency,
            pd.to_datetime(cutoff_date).strftime("%Y-%m-%d") if pd.notna(cutoff_date) else None,
            write_excel,
            output_path or self.config["outputs"]["excel_path"],
        )
        workload_df = processor.calculate_workload(cutoff_date=actual_cutoff_date)
        bow_cutoff_date = processor.infer_remaining_bow_cutoff_date()
        metric_definitions = pd.DataFrame(processor.output_metric_definitions)
        metric_lookup = metric_definitions.set_index("Metric").to_dict("index")

        control_df = pd.DataFrame(
            [
                {"Item": "Current Date", "Value": processor.get_current_cutoff_date()},
                {"Item": "Actual Cutoff Date", "Value": actual_cutoff_date},
                {"Item": "BoW Cutoff Date", "Value": bow_cutoff_date},
                {
                    "Item": "BoW Cutoff Definition",
                    "Value": "Later of current date and latest Original T0 in the Master File",
                },
                {"Item": "Output Frequency", "Value": processor.frequency},
                {
                    "Item": "Completion Input Frequency",
                    "Value": processor.input_completion_percentage_config["frequency"],
                },
                {"Item": "BoW Input Frequency", "Value": processor.input_bow_volume_config["frequency"]},
                {"Item": "WIP Received Status", "Value": ", ".join(sorted(processor.open_start_status))},
                {"Item": "Pending QC/BA Status", "Value": ", ".join(sorted(processor.pending_qc_ba_status))},
                {"Item": "Actual Start Status", "Value": ", ".join(sorted(processor.actual_start_status))},
                {
                    "Item": "Actual Completion Status",
                    "Value": ", ".join(sorted(processor.actual_completion_status)),
                },
                {"Item": "WBH Letter Rule", "Value": self.format_rule_days(processor.wbh_letter_days)},
                {"Item": "WBH Call Rule", "Value": self.format_rule_days(processor.wbh_call_days)},
                {"Item": "Completion UPT", "Value": processor.completion_upt},
                {"Item": "Init UPT", "Value": processor.init_upt},
                {"Item": "Working Hour", "Value": processor.working_hour},
                {
                    "Item": "Remaining BoW Plan Rule",
                    "Value": (
                        "Monthly: input BoW. Weekly/daily: max(input BoW - received through "
                        "actual cutoff at input frequency, 0). Allocate from actual cutoff "
                        "across current and future output periods."
                    ),
                },
                {
                    "Item": "Remaining BoW Output Rule",
                    "Value": "max(Output Plan Volume - Received Through BoW Cutoff, 0)",
                },
                {
                    "Item": "Remaining BoW Reconciliation",
                    "Value": (
                        "Output Plan Volume - Received Through BoW Cutoff - Remaining BoW Volume "
                        "+ Over Received Volume = 0"
                    ),
                },
                {
                    "Item": "Demand FTE Formula",
                    "Value": "(Completion Volume * Completion UPT + Init Volume * Init UPT) / Working Hour",
                },
            ]
        )

        metric_order = {
            metric: order
            for order, metric in enumerate(metric_definitions["Metric"])
        }
        workload_long_source = workload_df.reset_index().melt(
            id_vars=["Case Type", "Period"],
            var_name="Metric",
            value_name="Volume"
        )
        workload_long_source["Category"] = workload_long_source["Metric"].map(
            lambda metric: metric_lookup.get(metric, {}).get("Category", "Other")
        )
        workload_long_source["Subcategory"] = workload_long_source["Metric"].map(
            lambda metric: metric_lookup.get(metric, {}).get("Subcategory", "Other")
        )
        workload_long_source["Display Name"] = workload_long_source["Metric"].map(
            lambda metric: metric_lookup.get(metric, {}).get("Display Name", metric)
        )
        workload_long_source["Metric Order"] = workload_long_source["Metric"].map(metric_order).fillna(999)

        workload_wide = workload_long_source.pivot_table(
            index=["Category", "Subcategory", "Metric Order", "Metric", "Display Name", "Case Type"],
            columns="Period",
            values="Volume",
            aggfunc="sum",
            fill_value=0
        ).reset_index()
        period_columns = sorted(
            [column for column in workload_wide.columns if isinstance(column, pd.Period)]
        )
        workload_wide = workload_wide[
            ["Category", "Subcategory", "Metric Order", "Metric", "Display Name", "Case Type"]
            + period_columns
        ]
        workload_wide = workload_wide.sort_values(["Metric Order", "Case Type"]).drop(columns=["Metric Order"])
        workload_wide.columns = [
            str(column) if isinstance(column, pd.Period) else column
            for column in workload_wide.columns
        ]

        workload_long = workload_long_source.drop(columns=["Metric Order"])
        workload_long = self.add_period_display_columns(workload_long, period_columns=["Period"])
        workload_long = workload_long[
            [
                "Case Type",
                "Period",
                "Period Start",
                "Period End",
                "Category",
                "Subcategory",
                "Metric",
                "Display Name",
                "Volume",
            ]
        ]

        output_tables = {
            "00_Control": control_df,
            "01_Workload_Wide": workload_wide,
            "02_Workload_Long": workload_long,
            "03_Metric_Definitions": metric_definitions,
            "10_Input_BoW": self.output_table_from_series(
                processor.input_bow_volume,
                "Input BoW Volume",
                period_columns=["Input Period"]
            ),
            "11_Planned_Start": self.output_table_from_series(
                processor.bow_volume,
                "Planned Start Volume",
                period_columns=["Period"]
            ),
            "12_Start_Reconciliation": self.build_start_reconciliation_output_table(
                actual_cutoff_date=actual_cutoff_date,
                bow_cutoff_date=bow_cutoff_date,
            ),
            "13_WIP_Received_Start": self.output_table_from_series(
                processor.open_received_start_volume,
                "WIP Received Start Volume",
                period_columns=["Start Period"]
            ),
            "14_Pending_QC_BA": self.output_table_from_series(
                processor.pending_qc_ba_volume,
                "Pending QC/BA Volume",
                period_columns=["Start Period"]
            ),
            "20_Completion": self.output_table_from_dataframe(
                processor.completion_volume,
                period_columns=["Period"]
            ),
            "21_Completion_Distribution": self.build_completion_distribution_output_table(),
            "30_WBH_Letter": self.output_table_from_dataframe(
                processor.wbh_letter_volume,
                period_columns=["Period"]
            ),
            "31_WBH_Call": self.output_table_from_dataframe(
                processor.wbh_call_volume,
                period_columns=["Period"]
            ),
        }

        if processor.frequency == "M":
            output_tables.update(
                {
                    "40_Calc_Forecast_Completion": self.build_forecast_completion_calculation_table(),
                    "41_Calc_WBH_Letter": self.build_wbh_action_calculation_table("WBH Letter"),
                    "42_Calc_WBH_Call": self.build_wbh_action_calculation_table("WBH Call"),
                    "43_Calc_Demand_FTE": self.build_demand_fte_calculation_table(),
                }
            )

        self.output_excel_tables = output_tables
        setattr(processor, "output_excel_tables", output_tables)
        logger.info(
            "Excel output tables built | frequency=%s table_rows=%s",
            processor.frequency,
            {
                table_name: 0 if output_df is None else len(output_df)
                for table_name, output_df in output_tables.items()
            },
        )
        if write_excel or output_path is not None:
            self.write_output_excel(output_path=output_path, output_tables=output_tables)

        return self.output_excel_tables

    def write_output_excel(self, output_path=None, output_tables=None):
        if output_path is None:
            output_path = self.config["outputs"]["excel_path"]

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        if output_tables is None:
            output_tables = self.output_excel_tables

        if not output_tables:
            raise ValueError("No output tables available. Run build_output_excel_tables first.")

        logger.info(
            "Writing Excel output | path=%s sheets=%d",
            output_path,
            len(output_tables),
        )
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        category_fills = {
            "Start Pipeline": PatternFill("solid", fgColor="E2F0D9"),
            "Completion": PatternFill("solid", fgColor="D9EAF7"),
            "Demand": PatternFill("solid", fgColor="EADCF8"),
            "WBH Action": PatternFill("solid", fgColor="FCE4D6"),
        }

        if os.path.exists(output_path):
            try:
                with open(output_path, "a+b"):
                    pass
            except PermissionError:
                base_path, extension = os.path.splitext(output_path)
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"{base_path}_{timestamp}{extension}"
                logger.warning(
                    "Excel output path locked; writing timestamped file | path=%s",
                    output_path,
                )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            writer.book.calculation.calcMode = "auto"
            writer.book.calculation.fullCalcOnLoad = True
            writer.book.calculation.forceFullCalc = True

            for sheet_name, output_df in output_tables.items():
                safe_sheet_name = sheet_name[:31]
                if output_df is None or output_df.empty:
                    output_df = pd.DataFrame({"Message": ["No data"]})
                else:
                    output_df = output_df.copy()

                output_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                worksheet = writer.sheets[safe_sheet_name]
                worksheet.freeze_panes = "F2" if safe_sheet_name == "01_Workload_Wide" else "A2"
                worksheet.auto_filter.ref = worksheet.dimensions

                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                header_by_column = {
                    cell.column: str(cell.value)
                    for cell in worksheet[1]
                    if cell.value is not None
                }

                if safe_sheet_name == "12_Start_Reconciliation" and worksheet.max_row >= 2:
                    column_by_header = {
                        header: column
                        for column, header in header_by_column.items()
                    }
                    plan_letter = get_column_letter(column_by_header["Output Plan Volume"])
                    received_letter = get_column_letter(column_by_header["Received Through BoW Cutoff"])
                    remaining_letter = get_column_letter(column_by_header["Remaining BoW Volume"])
                    over_received_letter = get_column_letter(column_by_header["Over Received Volume"])
                    difference_letter = get_column_letter(column_by_header["Reconciliation Difference"])
                    status_letter = get_column_letter(column_by_header["Reconciliation Status"])

                    for row_number in range(2, worksheet.max_row + 1):
                        worksheet[f"{difference_letter}{row_number}"] = (
                            f"={plan_letter}{row_number}-{received_letter}{row_number}"
                            f"-{remaining_letter}{row_number}+{over_received_letter}{row_number}"
                        )
                        worksheet[f"{status_letter}{row_number}"] = (
                            f'=IF(ABS({difference_letter}{row_number})<0.000001,"OK","CHECK")'
                        )

                    status_range = f"{status_letter}2:{status_letter}{worksheet.max_row}"
                    worksheet.conditional_formatting.add(
                        status_range,
                        FormulaRule(
                            formula=[f'${status_letter}2="OK"'],
                            fill=PatternFill("solid", fgColor="E2F0D9"),
                        )
                    )
                    worksheet.conditional_formatting.add(
                        status_range,
                        FormulaRule(
                            formula=[f'${status_letter}2="CHECK"'],
                            fill=PatternFill("solid", fgColor="F4CCCC"),
                        )
                    )

                for row in worksheet.iter_rows(min_row=2):
                    category_value = None
                    metric_value = None
                    for cell in row:
                        if header_by_column.get(cell.column) == "Category":
                            category_value = cell.value
                        elif header_by_column.get(cell.column) == "Metric":
                            metric_value = cell.value
                        if category_value is not None and metric_value is not None:
                            break

                    if category_value in category_fills:
                        for cell in row:
                            cell.fill = category_fills[category_value]

                    for cell in row:
                        header = header_by_column.get(cell.column, "")
                        if metric_value == "Demand FTE" and isinstance(cell.value, (int, float)):
                            cell.number_format = "#,##0.00"
                        elif "FTE" in header and isinstance(cell.value, (int, float)):
                            cell.number_format = "#,##0.00"
                        elif "Probability" in header:
                            cell.number_format = "0.0%"
                        elif ("UPT" in header or "Hour" in header) and isinstance(cell.value, (int, float)):
                            cell.number_format = "#,##0.0"
                        elif "Volume" in header or header == "Value":
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0.0"
                        elif "Date" in header or header.endswith("Start") or header.endswith("End"):
                            cell.number_format = "yyyy-mm-dd"
                        elif isinstance(cell.value, (int, float)):
                            cell.number_format = "#,##0.0"

                        if (
                            safe_sheet_name == "00_Control" and header == "Value"
                        ) or (
                            safe_sheet_name == "03_Metric_Definitions"
                            and header in {"Source", "Logic", "Cutoff"}
                        ):
                            cell.alignment = Alignment(vertical="top", wrap_text=True)

                for column_cells in worksheet.columns:
                    column_letter = get_column_letter(column_cells[0].column)
                    max_length = 0
                    for cell in column_cells:
                        if cell.value is None:
                            continue
                        max_length = max(max_length, len(str(cell.value)))

                    worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)

                if safe_sheet_name == "00_Control":
                    worksheet.column_dimensions["A"].width = 34
                    worksheet.column_dimensions["B"].width = 80
                    for row_number in range(2, worksheet.max_row + 1):
                        if len(str(worksheet[f"B{row_number}"].value or "")) > 70:
                            worksheet.row_dimensions[row_number].height = 42
                elif safe_sheet_name == "03_Metric_Definitions":
                    definition_widths = {
                        "A": 16,
                        "B": 20,
                        "C": 30,
                        "D": 30,
                        "E": 42,
                        "F": 85,
                        "G": 52,
                    }
                    for column_letter, width in definition_widths.items():
                        worksheet.column_dimensions[column_letter].width = width
                    for row_number in range(2, worksheet.max_row + 1):
                        worksheet.row_dimensions[row_number].height = 45
                elif safe_sheet_name == "12_Start_Reconciliation":
                    reconciliation_widths = {
                        "A": 14,
                        "B": 14,
                        "C": 14,
                        "D": 14,
                        "E": 18,
                        "F": 18,
                        "G": 20,
                        "H": 30,
                        "I": 22,
                        "J": 22,
                        "K": 26,
                        "L": 22,
                    }
                    for column_letter, width in reconciliation_widths.items():
                        worksheet.column_dimensions[column_letter].width = width

        self.output_excel_path = output_path
        setattr(self.processor, "output_excel_path", output_path)
        logger.info(
            "Excel output written | path=%s size_bytes=%d",
            output_path,
            os.path.getsize(output_path) if os.path.exists(output_path) else 0,
        )
        return self.output_excel_path
```
