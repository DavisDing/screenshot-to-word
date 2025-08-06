import os
import sys

def get_base_path(logger=None):
    """
    获取程序运行目录（支持py和exe）
    """
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        if logger:
            logger.log(f"程序基础路径: {base_path}")
        return base_path
    except Exception as e:
        if logger:
            logger.log(f"获取基础路径异常: {e}")
        raise

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