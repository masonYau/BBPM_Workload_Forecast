import pandas as pd
from headers import MasterHeader as mh
from headers import BowHeader as bh
from headers import RamHeader as rh
import logging
import numpy as np

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
        }
        self.current_date = current_date
        self.frequency = frequency
        self.frequency_days = {"D": 1, "W": 7, "M": 30}
        self.completion_distribution = pd.Series()
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


    def calculate_workload(self):
        return



from datetime import datetime
current_time = pd.to_datetime(datetime.now())
self = DataProcessor(current_time, "W")
self.read_data()
self.calculate_completion_distribution()