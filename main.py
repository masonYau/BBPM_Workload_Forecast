import argparse
from datetime import datetime
import logging
import os
import sys
import time

import pandas as pd

from config_loader import load_config
from data_processor import DataProcessor
from workload_html_report import WorkloadHtmlVisualizer
from workload_excel_report import WorkloadExcelReporter


HTML_FREQUENCY_ORDER = ("M", "W", "D")
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_FILE_PATH = os.path.join(PROJECT_ROOT, "workload_forecast.log")
LOG_FORMAT = "%(asctime)s %(levelname)s [%(name)s] %(message)s"
CONSOLE_HANDLER_MARKER = "_workload_forecast_console_handler"
logger = logging.getLogger(__name__)


def setup_logging(log_path=None, level=logging.INFO):
    log_path = os.path.abspath(log_path or LOG_FILE_PATH)
    log_dir = os.path.dirname(log_path)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    formatter = logging.Formatter(LOG_FORMAT)

    has_log_handler = any(
        isinstance(handler, logging.FileHandler)
        and os.path.abspath(handler.baseFilename) == log_path
        for handler in root_logger.handlers
    )
    if not has_log_handler:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setLevel(level)
        handler.setFormatter(formatter)
        root_logger.addHandler(handler)

    has_console_handler = any(
        getattr(handler, CONSOLE_HANDLER_MARKER, False)
        for handler in root_logger.handlers
    )
    if not has_console_handler:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        setattr(console_handler, CONSOLE_HANDLER_MARKER, True)
        root_logger.addHandler(console_handler)

    return log_path


def get_html_frequencies(config):
    frequency_days = config["run"]["frequency_days"]
    configured_frequencies = list(frequency_days)

    frequencies = [
        frequency
        for frequency in HTML_FREQUENCY_ORDER
        if frequency in frequency_days
    ]
    frequencies.extend(
        frequency
        for frequency in configured_frequencies
        if frequency not in frequencies
    )
    return frequencies


def build_workload_processor(current_date, frequency, config=None):
    setup_logging()
    start_time = time.perf_counter()
    logger.info(
        "Processor start | frequency=%s current_date=%s",
        frequency,
        pd.to_datetime(current_date).strftime("%Y-%m-%d"),
    )
    processor = DataProcessor(pd.to_datetime(current_date), frequency, config=config)
    try:
        processor.run()
    except Exception:
        logger.error("Processor failed | frequency=%s", frequency, exc_info=True)
        raise

    period_count = 0
    case_type_count = 0
    if not processor.workload_volume.empty:
        period_count = processor.workload_volume.index.get_level_values("Period").nunique()
        case_type_count = processor.workload_volume.index.get_level_values("Case Type").nunique()
    logger.info(
        "Processor complete | frequency=%s elapsed_sec=%.2f workload_rows=%d periods=%d case_types=%d",
        frequency,
        time.perf_counter() - start_time,
        len(processor.workload_volume),
        period_count,
        case_type_count,
    )
    return processor


def run_workload_forecast(current_date=None, frequency=None, output_path=None, config_path=None):
    log_path = setup_logging()
    start_time = time.perf_counter()
    if current_date is None:
        current_date = datetime.now()

    config = load_config(config_path)
    if frequency is None:
        frequency = config["run"]["default_forecast_frequency"]

    logger.info(
        "Excel forecast run start | current_date=%s frequency=%s config_path=%s output_path=%s log_path=%s",
        pd.to_datetime(current_date).strftime("%Y-%m-%d"),
        frequency,
        config_path,
        output_path or config["outputs"]["excel_path"],
        log_path,
    )
    processor = build_workload_processor(current_date, frequency, config=config)
    reporter = WorkloadExcelReporter(processor)
    try:
        reporter.build_output_excel_tables(output_path=output_path)
    except Exception:
        logger.exception("Excel forecast run failed | frequency=%s", frequency)
        raise

    logger.info(
        "Excel forecast run complete | frequency=%s output_path=%s elapsed_sec=%.2f",
        frequency,
        reporter.output_excel_path,
        time.perf_counter() - start_time,
    )
    return processor, reporter


def run_workload_visualization(current_date=None, frequencies=None, output_path=None, config_path=None):
    log_path = setup_logging()
    start_time = time.perf_counter()
    if current_date is None:
        current_date = datetime.now()

    config = load_config(config_path)
    if frequencies is None:
        frequencies = get_html_frequencies(config)

    logger.info(
        "HTML visualization run start | current_date=%s frequencies=%s config_path=%s output_path=%s log_path=%s",
        pd.to_datetime(current_date).strftime("%Y-%m-%d"),
        ",".join(frequencies),
        config_path,
        output_path or config["outputs"]["html_path"],
        log_path,
    )
    processors_by_frequency = {
        frequency: build_workload_processor(current_date, frequency, config=config)
        for frequency in frequencies
    }
    visualizer = WorkloadHtmlVisualizer(processors_by_frequency, output_path=output_path)
    try:
        visualizer.write_html()
    except Exception:
        logger.exception("HTML visualization run failed | frequencies=%s", ",".join(frequencies))
        raise

    logger.info(
        "HTML visualization run complete | output_path=%s frequencies=%s elapsed_sec=%.2f",
        visualizer.output_path,
        ",".join(visualizer.visualization_data.get("frequencies", {}).keys()),
        time.perf_counter() - start_time,
    )
    return visualizer


def main(config_path=None):
    log_path = setup_logging()
    start_time = time.perf_counter()
    config = load_config(config_path)
    run_date = datetime.now()
    default_frequency = config["run"]["default_forecast_frequency"]
    html_frequencies = get_html_frequencies(config)
    logger.info(
        "Workload forecast main start | current_date=%s config_path=%s default_frequency=%s html_frequencies=%s log_path=%s",
        pd.to_datetime(run_date).strftime("%Y-%m-%d"),
        config_path,
        default_frequency,
        ",".join(html_frequencies),
        log_path,
    )

    try:
        default_processor = build_workload_processor(run_date, default_frequency, config=config)

        excel_reporter = WorkloadExcelReporter(default_processor)
        excel_reporter.build_output_excel_tables()
        print(f"Output Excel written to: {excel_reporter.output_excel_path}")

        processors = {default_frequency: default_processor}
        for frequency in html_frequencies:
            if frequency not in processors:
                processors[frequency] = build_workload_processor(run_date, frequency, config=config)

        html_visualizer = WorkloadHtmlVisualizer(processors)
        html_visualizer.write_html()
        print(f"Output HTML written to: {html_visualizer.output_path}")
    except Exception:
        logger.exception("Workload forecast main failed")
        raise

    logger.info(
        "Workload forecast main complete | excel_path=%s html_path=%s elapsed_sec=%.2f",
        excel_reporter.output_excel_path,
        html_visualizer.output_path,
        time.perf_counter() - start_time,
    )


def parse_args():
    parser = argparse.ArgumentParser(description="Generate workload forecast Excel and HTML reports.")
    parser.add_argument(
        "config_path",
        nargs="?",
        help="Optional path to config.json.",
    )
    parser.add_argument(
        "--config",
        dest="config_path_option",
        metavar="CONFIG_PATH",
        help="Optional path to config.json. Overrides the positional config path when both are provided.",
    )
    args = parser.parse_args()
    return args.config_path_option or args.config_path


if __name__ == "__main__":
    main(config_path=parse_args())
