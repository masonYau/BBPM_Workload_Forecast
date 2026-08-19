# config.json

```json
{
  "inputs": {
    "completion_percentage": {
      "file_path": "Input_Completion_Percentage.xlsx",
      "sheet_names": null,
      "period_column": "Month",
      "percentage_column": "Percentage",
      "frequency": "M"
    },
    "bow_volume": {
      "file_path": "Input_BoW_Volume.xlsx",
      "sheet_names": null,
      "period_column": "Month",
      "volume_column": "Volume",
      "frequency": "M"
    },
    "tracker": {
      "file_path": "BBPM Case tracker*.xlsm",
      "sheet_name": "Master File",
      "skiprows": 1
    }
  },
  "run": {
    "default_forecast_frequency": "M",
    "visualization_frequencies": ["M", "W", "D"],
    "frequency_days": {
      "D": 1,
      "W": 7,
      "M": 30
    }
  },
  "business_rules": {
    "completed_source_statuses": ["E2E Completed", "ReCompleted - WBH"],
    "status_mapping": {
      "CSEM": "Completed",
      "Cancelled - 2nd AC Opening under NTB": "Cancelled/Closed",
      "Cancelled - AC re-opened under NTB": "Cancelled/Closed",
      "Cancelled - All AC Closed": "Completed",
      "Cancelled - CDAB Code": "Cancelled/Closed",
      "Cancelled - China Notch-Down": "Cancelled/Closed",
      "Cancelled - KPMG Managed": "Cancelled/Closed",
      "Cancelled - RM Managed": "Cancelled/Closed",
      "Cancelled - trigger reason can be discounted": "Cancelled/Closed",
      "Cat C CSEM - CDD deficiencies (Geographical risk)": "Completed",
      "Cat C CSEM - Tax Risk Appetite": "Completed",
      "COMPLETED": "WIP",
      "E2E Completed": "Completed",
      "FAILED": "WIP",
      "IN_PROGRESS": "WIP",
      "Not Loading": "Not Initiated",
      "Pending BA Approval": "Pending QC/BA",
      "Pending Data Capture": "WIP",
      "Pending QC": "Pending QC/BA",
      "WBH Imposed": "WBH",
      "ReCompleted - WBH": "Completed",
      "Completed - WBH with KYCH over 30 days": "Completed"
    },
    "actual_start_statuses": ["Completed", "WBH", "Cancelled/Closed", "WIP", "Pending QC/BA"],
    "open_start_statuses": ["WBH", "WIP", "Not Initiated", "Pending QC/BA"],
    "pending_qc_ba_statuses": ["Pending QC/BA"],
    "actual_completion_statuses": ["Completed"],
    "wbh_letter_days": {
      "PR": 90,
      "Trigger": 60
    },
    "wbh_call_days": {
      "PR": 95,
      "Trigger": 65
    },
    "completion_upt": 4.5,
    "init_upt": 0.5,
    "working_hour": 129,
    "read_completion_percentage_from_input": true
  },
  "outputs": {
    "excel_path": "data/Workload_Forecast_Output.xlsx",
    "html_path": "data/Workload_Forecast_Visualization.html",
    "log_path": "./mi.log"
  }
}
```
