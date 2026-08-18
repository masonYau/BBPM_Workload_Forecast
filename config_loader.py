import json
import os


CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")


def load_config(config_path=None):
    path = config_path or CONFIG_PATH
    with open(path, "r", encoding="utf-8") as config_file:
        return json.load(config_file)
