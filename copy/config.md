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
      "E2E Completed": "Completed",
      "WBH Imposed": "WBH",
      "ReCompleted - WBH": "Completed",
      "Cancelled - All AC Closed": "Completed",
      "CSEM": "Completed",
      "Cancelled - KPMG Managed": "Cancelled",
      "Cancelled - RM Managed": "Cancelled",
      "IN_PROGRESS": "WIP",
      "Cancelled - 2nd AC Opening under NTB": "Cancelled",
      "Cancelled - AC re-opened under NTB": "Cancelled",
      "Pending BA Approval": "WIP",
      "Not Loading": "Not Initiated"
    },
    "actual_start_statuses": ["Completed", "WBH", "Cancelled", "WIP"],
    "open_start_statuses": ["WBH", "WIP", "Not Initiated"],
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
