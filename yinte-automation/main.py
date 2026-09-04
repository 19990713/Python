#!/usr/bin/env python
"""
印特 GUI 自动化脚本 — 从 Excel 读取数据自动录入

使用方法:
    py -3.12 main.py --excel "数据.xlsx"
    py -3.12 main.py --excel "数据.xlsx" --dry-run    # 试运行
    py -3.12 main.py --excel "数据.xlsx" --start-row 3 # 从第3行开始

前置条件:
    1. 印特已打开，"新建工单"窗口可见
    2. Excel 列: 尺寸分类, 子项, 数量, 份数, 页数
"""

import argparse
import sys
from pathlib import Path

import yaml

# 脚本所在目录
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))

from src.logger import setup_logging
from src.excel_reader import read_excel, read_cell
from src.automation import AutomationEngine


def load_config(config_path: str) -> dict:
    """加载 config.yaml，支持 PyInstaller 冻结模式。"""
    path = Path(config_path)
    if not path.exists():
        # PyInstaller 冻结模式：尝试从 sys._MEIPASS 读取内嵌配置
        if getattr(sys, 'frozen', False):
            import os
            bundled = Path(sys._MEIPASS) / "config.yaml"
            if bundled.exists():
                with open(bundled, "r", encoding="utf-8") as f:
                    return yaml.safe_load(f)
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# 有效图纸类型
VALID_SUBTYPES = {"蓝图", "彩图", "白图", "复图", "硫酸纸"}


def main():
    parser = argparse.ArgumentParser(description="印特 GUI 自动化录入")
    parser.add_argument("--excel", "-e", required=True, help="Excel 文件路径")
    parser.add_argument("--config", "-c", default=str(SCRIPT_DIR / "config.yaml"), help="配置文件")
    parser.add_argument("--sheet", "-s", default=0, help="工作表名或索引")
    parser.add_argument("--dry-run", action="store_true", help="试运行")
    parser.add_argument("--start-row", type=int, default=0, help="起始行(0=第一行)")
    args = parser.parse_args()

    logger = setup_logging(str(SCRIPT_DIR / "logs"))
    config = load_config(args.config)

    # 读取 Excel
    excel_cfg = config.get("excel", {})
    header_row = excel_cfg.get("header_row", 0)  # 0 = 自动检测
    logger.info(f"读取: {args.excel} (标题行: {header_row})")
    try:
        sheet = int(args.sheet)
    except ValueError:
        sheet = args.sheet
    records = read_excel(args.excel, sheet_name=sheet, header_row=header_row)

    # 兼容旧格式: "A1加长" → "A1+"
    for r in records:
        val = r.get("尺寸分类", "")
        if isinstance(val, str) and "加长" in val:
            r["尺寸分类"] = val.replace("加长", "+")

    if args.start_row > 0:
        records = records[args.start_row:]

    if not records:
        logger.warning("无数据")
        sys.exit(0)

    # 获取列配置
    col_cfg = config.get("columns", {})

    # 过滤掉汇总行
    size_col = col_cfg["尺寸分类"]
    records = [r for r in records if r.get(size_col, "").strip() not in ("总计", "合计", "小计", "")]
    if not records:
        logger.warning("过滤后无数据")
        sys.exit(0)

    # 校验必要列 (只检查尺寸分类、份数、页数；数量列可选)
    required = [col_cfg["尺寸分类"], col_cfg["份数"], col_cfg["页数"]]
    if col_cfg.get("子项"):
        required.append(col_cfg["子项"])
    excel_cols = set(records[0].keys())
    missing = [r for r in required if r not in excel_cols]
    if missing:
        logger.error(f"Excel缺少列: {missing}")
        logger.info(f"实际列: {sorted(excel_cols)}")
        logger.info("请在 config.yaml 的 excel.header_row 设置正确的标题行号")
        sys.exit(1)

    logger.info(f"共 {len(records)} 条数据")

    # 自动读取子项类型 (从 B4 单元格)
    subtype_cell = config.get("subtype_cell", "B4")
    cell_value = read_cell(args.excel, sheet, subtype_cell)
    if cell_value:
        if cell_value in VALID_SUBTYPES:
            config["default_subtype"] = cell_value
            logger.info(f"从 {subtype_cell} 读取到子项: {cell_value}")
        else:
            logger.warning(f"{subtype_cell} 内容 '{cell_value}' 不是有效图纸类型，使用默认: {config.get('default_subtype', '蓝图')}")
    else:
        config["default_subtype"] = config.get("default_subtype", "蓝图")
    logger.info(f"子项: {config['default_subtype']}")

    if args.dry_run:
        logger.info("=== 试运行 ===")
        default_sub = config.get("default_subtype", "")
        for i, row in enumerate(records):
            size = row.get(col_cfg["尺寸分类"], "")
            sub = row.get(col_cfg.get("子项", ""), "") or default_sub
            code = config["product_codes"].get(f"{sub}_{size}", config["product_codes"].get(size, "???"))
            logger.info(
                f"  {i + 1}. {size}/{sub} → {code} | "
                f"数量={row.get(col_cfg['数量'])} "
                f"份数={row.get(col_cfg['份数'])} "
                f"页数={row.get(col_cfg['页数'])}"
            )
        sys.exit(0)

    # 运行
    engine = AutomationEngine(config)
    engine.connect()
    result = engine.run(records)

    logger.info(f"成功 {result['success']}, 失败 {result['failed']}")
    sys.exit(2 if result["failed"] > 0 else 0)


if __name__ == "__main__":
    main()
