import json
import os
from typing import Any, Optional

from util import system

user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0"

DEBUG = False

global_config: Optional[dict[str, Any]] = None


def get_config_path():
    return os.path.join(system.get_exe_dir(), "config.json")


def get_config(throw_exception=True) -> dict[str, Any]:
    global global_config

    if global_config is not None:
        return global_config

    try:
        with open(get_config_path(), "r") as f:
            global_config = json.load(f)

        return global_config
    except Exception as e:
        if throw_exception:
            raise e
        else:
            return {}


def save_config(config_data) -> Optional[Exception]:
    global global_config

    try:
        with open(get_config_path(), "w", encoding="utf-8") as f:
            f.write(json.dumps(config_data))

        global_config = config_data
        return None
    except Exception as e:
        return e
