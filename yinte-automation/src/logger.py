"""
日志模块 — 同时输出到文件和控制台
"""

import logging
import sys
from pathlib import Path
from datetime import datetime


def setup_logging(log_dir: str = "logs") -> logging.Logger:
    """配置日志，返回 root logger。"""
    log_path = Path(log_dir)
    log_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_path / f"yinte_automation_{timestamp}.log"

    # 日志格式
    file_fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    console_fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
    )

    logger = logging.getLogger("yinte")
    logger.setLevel(logging.DEBUG)

    # 文件 handler — 记录所有级别
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(file_fmt)
    logger.addHandler(fh)

    # 控制台 handler — 只显示 INFO 以上
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(console_fmt)
    logger.addHandler(ch)

    logger.info(f"日志文件: {log_file}")
    return logger
