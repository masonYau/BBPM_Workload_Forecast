from datetime import datetime

import pandas as pd

from data_processor import DataProcessor
from workload_excel_report import WorkloadExcelReporter


def run_workload_forecast(current_date=None, frequency="M", output_path=None):
    if current_date is None:
        current_date = datetime.now()

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

    reporter = WorkloadExcelReporter(processor)
    reporter.build_output_excel_tables(output_path=output_path)
    return processor, reporter


if __name__ == "__main__":
    _, excel_reporter = run_workload_forecast()
    print(f"Output Excel written to: {excel_reporter.output_excel_path}")
