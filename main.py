from datetime import datetime

import pandas as pd

from config_loader import load_config
from data_processor import DataProcessor
from workload_html_report import WorkloadHtmlVisualizer
from workload_excel_report import WorkloadExcelReporter


def build_workload_processor(current_date, frequency, config=None):
    processor = DataProcessor(pd.to_datetime(current_date), frequency, config=config)
    processor.run()
    return processor


def run_workload_forecast(current_date=None, frequency=None, output_path=None):
    if current_date is None:
        current_date = datetime.now()

    config = load_config()
    if frequency is None:
        frequency = config["run"]["default_forecast_frequency"]

    processor = build_workload_processor(current_date, frequency, config=config)
    reporter = WorkloadExcelReporter(processor)
    reporter.build_output_excel_tables(output_path=output_path)
    return processor, reporter


def run_workload_visualization(current_date=None, frequencies=None, output_path=None):
    if current_date is None:
        current_date = datetime.now()

    config = load_config()
    if frequencies is None:
        frequencies = config["run"]["visualization_frequencies"]

    processors_by_frequency = {
        frequency: build_workload_processor(current_date, frequency, config=config)
        for frequency in frequencies
    }
    visualizer = WorkloadHtmlVisualizer(processors_by_frequency, output_path=output_path)
    visualizer.write_html()
    return visualizer


def main():
    config = load_config()
    run_date = datetime.now()
    default_frequency = config["run"]["default_forecast_frequency"]
    default_processor = build_workload_processor(run_date, default_frequency, config=config)

    excel_reporter = WorkloadExcelReporter(default_processor)
    excel_reporter.build_output_excel_tables()
    print(f"Output Excel written to: {excel_reporter.output_excel_path}")

    processors = {default_frequency: default_processor}
    for frequency in config["run"]["visualization_frequencies"]:
        if frequency not in processors:
            processors[frequency] = build_workload_processor(run_date, frequency, config=config)

    html_visualizer = WorkloadHtmlVisualizer(processors)
    html_visualizer.write_html()
    print(f"Output HTML written to: {html_visualizer.output_path}")


if __name__ == "__main__":
    main()
