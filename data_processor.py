import pandas as pd
from headers import MasterHeader as mh
from headers import BowHeader as bh
from headers import RamHeader as rh
import logging
import numpy as np
import math

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

        self.input_completion_percentage_config = {
            "file_path": "Input_Completion_Percentage.xlsx",
            "sheet_name": "Sheet1",
            "period_column": "Month",
            "percentage_column": "Percentage",
            "frequency": "M",
        }

    def read_data(self):

        master_df = pd.read_excel("BBPM Case tracker.xlsm", sheet_name="Master File", skiprows=1)
        master_df[mh.CIN] = master_df[mh.CIN].astype(str).apply(lambda x: x.lstrip('0'))
        master_df[mh.OriginalT0] = pd.to_datetime(master_df[mh.OriginalT0])
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
        input_df = pd.read_excel(config["file_path"], sheet_name=config["sheet_name"])
        input_df = input_df[[config["period_column"], config["percentage_column"]]].dropna()
        input_df[config["period_column"]] = pd.to_numeric(input_df[config["period_column"]], errors="coerce")
        input_df[config["percentage_column"]] = pd.to_numeric(input_df[config["percentage_column"]], errors="coerce")
        input_df = input_df.dropna()

        completion_distribution = input_df.groupby(config["period_column"])[config["percentage_column"]].sum()
        completion_distribution.index = completion_distribution.index.astype(int)
        if len(completion_distribution) > 0 and completion_distribution.max() > 1:
            completion_distribution = completion_distribution / 100

        self.completion_distribution = self.convert_completion_distribution_frequency(
            completion_distribution,
            input_frequency=config["frequency"]
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


    def calculate_workload(self):
        return



from datetime import datetime
current_time = pd.to_datetime(datetime.now())
self = DataProcessor(current_time, "W")
# self.read_data()
# self.calculate_completion_distribution()
self.read_input_completion_percentage()