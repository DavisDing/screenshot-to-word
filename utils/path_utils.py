import os
import sys


def get_base_path():
    """Return executable directory when packaged, otherwise the project root."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def ensure_directories(base_path, logger=None):
    required_dirs = ["excel_input", "word_output", "logs", "Temp"]
    for folder in required_dirs:
        path = os.path.join(base_path, folder)
        if not os.path.exists(path):
            os.makedirs(path)
            if logger:
                logger.log(f"创建目录：{path}")
        else:
            if logger:
                logger.log(f"目录已存在：{path}")
