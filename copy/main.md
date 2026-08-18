# main.py

```python
from datetime import datetime

import pandas as pd

from data_processor import DataProcessor
from workload_html_report import WorkloadHtmlVisualizer
from workload_excel_report import WorkloadExcelReporter


def build_workload_processor(current_date, frequency):
    processor = DataProcessor(pd.to_datetime(current_date), frequency)
    processor.read_data()
    processor.calculate_completion_distribution()
    processor.read_input_completion_percentage()
    processor.read_input_bow_volume()
    processor.calculate_remaining_bow_volume()
    processor.calculate_actual_start_volume()
    processor.calculate_completion_volume()
    processor.calculate_wbh_letter_volume()
    processor.calculate_wbh_call_volume()
    processor.calculate_workload()
    return processor


def run_workload_forecast(current_date=None, frequency="M", output_path=None):
    if current_date is None:
        current_date = datetime.now()

    processor = build_workload_processor(current_date, frequency)
    reporter = WorkloadExcelReporter(processor)
    reporter.build_output_excel_tables(output_path=output_path)
    return processor, reporter


def run_workload_visualization(current_date=None, frequencies=("M", "W", "D"), output_path=None):
    if current_date is None:
        current_date = datetime.now()

    processors_by_frequency = {
        frequency: build_workload_processor(current_date, frequency)
        for frequency in frequencies
    }
    visualizer = WorkloadHtmlVisualizer(processors_by_frequency, output_path=output_path)
    visualizer.write_html()
    return visualizer


if __name__ == "__main__":
    run_date = datetime.now()
    monthly_processor = build_workload_processor(run_date, "M")

    excel_reporter = WorkloadExcelReporter(monthly_processor)
    excel_reporter.build_output_excel_tables()
    print(f"Output Excel written to: {excel_reporter.output_excel_path}")

    processors = {
        "M": monthly_processor,
        "W": build_workload_processor(run_date, "W"),
        "D": build_workload_processor(run_date, "D"),
    }
    html_visualizer = WorkloadHtmlVisualizer(processors)
    html_visualizer.write_html()
    print(f"Output HTML written to: {html_visualizer.output_path}")
```
