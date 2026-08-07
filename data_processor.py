import pandas as pd
from headers import MasterHeader as mh
from headers import BowHeader as bh
from headers import RamHeader as rh
import logging
import numpy as np
import math
import os

class DataProcessor:

    def __init__(self, current_date: pd.Timestamp, frequency: str):

        self.master_df = pd.DataFrame()
        self.comple_status = {'E2E Completed', 'ReCompleted - WBH'}
        self.wbh_status = {
            'E2E Completed': 'Completed',
            'WBH Imposed': 'WBH',
            'ReCompleted - WBH': 'Completed',
            'Cancelled - All AC Closed': 'Completed',
            'CSEM': 'Completed',
            'Cancelled - KPMG Managed': 'Cancelled',
            'Cancelled - RM Managed': 'Cancelled',
            'IN_PROGRESS': 'WIP',
            'Cancelled - 2nd AC Opening under NTB': 'Cancelled',
            'Cancelled - AC re-opened under NTB': 'Cancelled',
            'Pending BA Approval': 'WIP',
            'Not Loading': 'Not Initiated'
        }
        self.current_date = current_date
        self.frequency = frequency
        self.frequency_days = {"D": 1, "W": 7, "M": 30}
        self.completion_distribution = pd.Series()
        self.input_bow_volume = pd.Series()
        self.bow_volume = pd.Series()
        self.received_volume = pd.Series()
        self.actual_start_volume = pd.Series()
        self.open_received_start_volume = pd.Series()
        self.forecast_completion_volume = pd.Series()
        self.actual_completion_volume = pd.Series()
        self.forecast_wbh_letter_volume = pd.Series()
        self.completion_volume = pd.DataFrame()
        self.wbh_letter_volume = pd.DataFrame()
        self.remaining_bow_volume = pd.DataFrame()
        self.actual_cutoff_date = pd.NaT
        self.remaining_bow_cutoff_date = pd.NaT
        self.actual_start_status = {'Completed', 'WBH', 'Cancelled', 'WIP'}
        self.open_start_status = {'WBH', 'WIP', 'Not Initiated'}
        self.actual_completion_status = {'Completed'}
        self.wbh_letter_days = {
            "PR": 90,
            "Trigger": 60,
        }

        self.input_completion_percentage_config = {
            "file_path": "Input_Completion_Percentage.xlsx",
            "sheet_names": None,
            "period_column": "Month",
            "percentage_column": "Percentage",
            "frequency": "M",
        }

        self.input_bow_volume_config = {
            "file_path": "Input_BoW_Volume.xlsx",
            "sheet_names": None,
            "period_column": "Month",
            "volume_column": "Volume",
            "frequency": "M",
        }

    def read_data(self):

        tracker_path = "BBPM Case tracker.xlsm"
        if not os.path.exists(tracker_path) and os.path.exists("data/BBPM Case tracker.xlsm"):
            tracker_path = "data/BBPM Case tracker.xlsm"

        master_df = pd.read_excel(tracker_path, sheet_name="Master File", skiprows=1)
        master_df[mh.CIN] = master_df[mh.CIN].astype(str).apply(lambda x: x.lstrip('0'))
        master_df[mh.OriginalT0] = pd.to_datetime(master_df[mh.OriginalT0])
        master_df[mh.ReviewType] = master_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in x else "Trigger")
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

    def read_input_completion_percentage(self):
        config = self.input_completion_percentage_config
        completion_distributions = {}

        for case_type, input_df in self.read_case_type_input_sheets(config):
            input_df = input_df[[config["period_column"], config["percentage_column"]]].dropna()
            input_df[config["period_column"]] = pd.to_numeric(input_df[config["period_column"]], errors="coerce")
            input_df[config["percentage_column"]] = pd.to_numeric(input_df[config["percentage_column"]], errors="coerce")
            input_df = input_df.dropna()

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

        for case_type, input_df in self.read_case_type_input_sheets(config):
            input_df = input_df[[config["period_column"], config["volume_column"]]].dropna()
            input_df[config["period_column"]] = pd.to_datetime(input_df[config["period_column"]], errors="coerce")
            input_df[config["volume_column"]] = pd.to_numeric(input_df[config["volume_column"]], errors="coerce")
            input_df = input_df.dropna()

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

    def read_case_type_input_sheets(self, config):
        file_path = self.resolve_input_file_path(config["file_path"])
        excel_file = pd.ExcelFile(file_path)
        sheet_names = config.get("sheet_names")
        if sheet_names is None:
            sheet_names = excel_file.sheet_names
        elif isinstance(sheet_names, str):
            sheet_names = [sheet_names]

        for sheet_name in sheet_names:
            input_df = pd.read_excel(file_path, sheet_name=sheet_name)
            required_columns = {config["period_column"]}
            if "percentage_column" in config:
                required_columns.add(config["percentage_column"])
            if "volume_column" in config:
                required_columns.add(config["volume_column"])

            if not required_columns.issubset(input_df.columns):
                continue

            input_df = input_df.dropna(how="all")
            if input_df.empty:
                continue

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
            self.remaining_bow_cutoff_date = min(original_t0.max().normalize(), current_date)

        return self.remaining_bow_cutoff_date

    def calculate_received_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

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
        return self.actual_start_volume

    def calculate_open_received_start_volume(self, cutoff_date=None):
        if cutoff_date is None:
            cutoff_date = self.infer_actual_cutoff_date()

        if mh.OriginalT0 not in self.master_df.columns or mh.ReviewType not in self.master_df.columns or mh.TaskStatus not in self.master_df.columns:
            self.open_received_start_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Start Period"])
            )
            self.open_received_start_volume.name = "Open Received Start Volume"
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

        self.open_received_start_volume.name = "Open Received Start Volume"
        return self.open_received_start_volume

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

        letter_date = start_period.start_time.normalize() + pd.Timedelta(days=letter_days)
        letter_period = letter_date.to_period(self.frequency)
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
                output_period = max(letter_period, cutoff_period)
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

        open_received_start_volume = self.calculate_open_received_start_volume(cutoff_date=cutoff_date)
        wbh_letter_volume = {}

        for (case_type, start_period), start_volume in open_received_start_volume.items():
            self.add_wbh_letter_volume(
                wbh_letter_volume,
                case_type,
                start_period,
                start_volume,
                source="Open Received",
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
        wbh_letter_df["Open Received WBH Letter Volume"] = source_volume.get("Open Received", 0)
        wbh_letter_df["Remaining BoW WBH Letter Volume"] = source_volume.get("Remaining BoW", 0)
        wbh_letter_df["Forecast WBH Letter Volume"] = (
            wbh_letter_df["Open Received WBH Letter Volume"]
            + wbh_letter_df["Remaining BoW WBH Letter Volume"]
        )

        self.forecast_wbh_letter_volume = wbh_letter_df["Forecast WBH Letter Volume"]
        self.forecast_wbh_letter_volume.name = "Forecast WBH Letter Volume"
        self.wbh_letter_volume = wbh_letter_df
        return self.wbh_letter_volume

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

        open_received_start_volume = self.calculate_open_received_start_volume(cutoff_date=cutoff_date)
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
        return self.completion_volume

    def calculate_remaining_bow_volume(self):
        if self.input_bow_volume.empty:
            self.read_input_bow_volume()

        remaining_cutoff_date = self.infer_remaining_bow_cutoff_date()
        self.calculate_received_volume(cutoff_date=remaining_cutoff_date)

        input_frequency = self.input_bow_volume_config["frequency"]
        if mh.OriginalT0 in self.master_df.columns and mh.ReviewType in self.master_df.columns:
            received_df = self.master_df[[mh.OriginalT0, mh.ReviewType]].copy()
            received_df[mh.ReviewType] = received_df[mh.ReviewType].apply(lambda x: "PR" if "PR" in str(x) else "Trigger")
            received_df[mh.OriginalT0] = pd.to_datetime(received_df[mh.OriginalT0], errors="coerce").dt.normalize()
            received_df = received_df.dropna(subset=[mh.OriginalT0])
            received_df = received_df[received_df[mh.OriginalT0] <= remaining_cutoff_date]
            received_df["Input Period"] = received_df[mh.OriginalT0].dt.to_period(input_frequency)
            received_by_input_period = received_df.groupby([mh.ReviewType, "Input Period"]).size()
            received_by_input_period.index = received_by_input_period.index.set_names(["Case Type", "Input Period"])
        else:
            received_by_input_period = pd.Series(dtype=int)

        remaining_output_volume = {}
        remaining_start_limit = remaining_cutoff_date + pd.Timedelta(days=1)

        for (case_type, input_period), planned_volume in self.input_bow_volume.sort_index().items():
            input_start_date = input_period.start_time.normalize()
            input_end_date = (input_period + 1).start_time.normalize()
            received_volume = received_by_input_period.get((case_type, input_period), 0)
            remaining_volume = max(planned_volume - received_volume, 0)
            remaining_start_date = max(input_start_date, remaining_start_limit)

            if remaining_volume <= 0 or remaining_start_date >= input_end_date:
                continue

            remaining_days = (input_end_date - remaining_start_date).days
            first_output_period = remaining_start_date.to_period(self.frequency)
            last_output_period = (input_end_date - pd.Timedelta(days=1)).to_period(self.frequency)

            for output_period in pd.period_range(first_output_period, last_output_period, freq=self.frequency):
                output_start_date = output_period.start_time.normalize()
                output_end_date = (output_period + 1).start_time.normalize()
                overlap_days = (
                    min(input_end_date, output_end_date) - max(remaining_start_date, output_start_date)
                ).days

                if overlap_days > 0:
                    remaining_output_volume[(case_type, output_period)] = (
                        remaining_output_volume.get((case_type, output_period), 0)
                        + remaining_volume * overlap_days / remaining_days
                    )

        if remaining_output_volume:
            remaining_output_volume = pd.Series(remaining_output_volume, dtype=float).sort_index()
            remaining_output_volume.index = pd.MultiIndex.from_tuples(
                remaining_output_volume.index,
                names=["Case Type", "Period"]
            )
        else:
            remaining_output_volume = pd.Series(
                dtype=float,
                index=pd.MultiIndex.from_arrays([[], []], names=["Case Type", "Period"])
            )
        remaining_output_volume.name = "Remaining BoW Volume"

        periods = self.bow_volume.index.union(remaining_output_volume.index).sort_values()
        remaining_df = pd.DataFrame(index=periods)
        remaining_df.index = remaining_df.index.set_names(["Case Type", "Period"])
        remaining_df["Input BoW Volume"] = self.bow_volume.reindex(periods, fill_value=0)
        remaining_df["Received Volume"] = self.received_volume.reindex(periods, fill_value=0)
        remaining_df["Remaining BoW Volume"] = remaining_output_volume.reindex(periods, fill_value=0)
        remaining_df["Over Received Volume"] = (
            remaining_df["Received Volume"] - remaining_df["Input BoW Volume"]
        ).clip(lower=0)

        self.remaining_bow_volume = remaining_df
        return self.remaining_bow_volume


    def calculate_workload(self):
        return



from datetime import datetime
current_time = pd.to_datetime(datetime.now())
self = DataProcessor(current_time, "W")
self.read_data()
self.calculate_completion_distribution()
self.read_input_completion_percentage()
self.read_input_bow_volume()
self.calculate_remaining_bow_volume()
self.calculate_actual_start_volume()
self.calculate_completion_volume()
self.calculate_wbh_letter_volume()
