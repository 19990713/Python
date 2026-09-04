"""
印特 GUI 自动化引擎 — 最终版
流程:
  第1行: 单击编号 → 代码 → Enter → skip → 份数 → Enter → 页数 → Enter
  第N行: Enter → 代码 → Enter → skip → 份数 → Enter → 页数 → Enter
补录模式:
  找到已打开的工单窗口 → 最大化 → 点击NewItem Row → 从当前位置开始录入
"""

import ctypes
import time
import logging
from typing import Dict, List, Any, Optional

from pywinauto import Application, keyboard, mouse
import pyperclip

logger = logging.getLogger("yinte.automation")


class AutomationEngine:
    """印特 GUI 自动化引擎。"""

    def __init__(self, config: Dict):
        self.config = config
        self.app_cfg = config.get("app", {})
        self.window_title = self.app_cfg.get("window_title", "新建工单")
        self.action_delay = self.app_cfg.get("action_delay", 0.3)
        # 补录模式额外延时：补录时工单已有数据，放慢节奏避免漏录/串行
        self.append_delay = self.app_cfg.get("append_delay", 0.4)
        self._extra_delay = 0.0

        # 相对坐标 (仅第1行使用)
        self.rel_x = self.app_cfg.get("rel_x", 111)
        self.rel_y_base = self.app_cfg.get("rel_y_base", 375)

        self.app: Optional[Application] = None
        self.window = None

    def connect(self):
        """连接到新建工单窗口。"""
        logger.info(f"连接窗口: '{self.window_title}'")
        self.app = Application(backend="win32").connect(
            title_re=f".*{self.window_title}.*"
        )
        self.window = self.app.window(title_re=f".*{self.window_title}.*")
        self.window.set_focus()
        time.sleep(0.3)
        logger.info(f"已连接: '{self.window.window_text()}'")

    def open_new_work_order(self) -> bool:
        """
        在印特主窗口中双击新建工单按钮。
        返回 True 表示成功打开。
        """
        from pywinauto import mouse as pmouse
        from pywinauto.findwindows import find_windows, find_elements

        logger.info("正在查找印特主窗口...")
        main_w = None
        try:
            # 策略1: 用标题匹配 (win32后端中文编码可能失败)
            try:
                main_app = Application(backend="win32").connect(title_re=".*印特.*")
                main_w = main_app.window(title_re=".*印特.*")
                logger.info("通过标题找到印特主窗口")
            except Exception:
                pass

            # 策略2: 用窗口类名匹配 (印特是Windows Forms程序)
            if main_w is None:
                els = find_elements(class_name_re="WindowsForms10.*Window\.8\.app.*")
                for el in els:
                    try:
                        t = el.name
                        # 排除 新建工单 窗口
                        if "新建工单" not in t:
                            main_app = Application(backend="win32").connect(handle=el.handle)
                            main_w = main_app.window(handle=el.handle)
                            logger.info(f"通过类名找到印特主窗口: {t}")
                            break
                    except Exception:
                        continue

            if main_w is None:
                logger.error("未找到印特主窗口，请确认印特已启动")
                return False
        except Exception:
            logger.error("未找到印特主窗口，请确认印特已启动")
            return False

        main_w.set_focus()
        time.sleep(0.3)

        wr = main_w.rectangle()
        # 新建工单按钮相对主窗口位置
        btn_x = wr.left + 128
        btn_y = wr.top + 79

        logger.info(f"双击新建工单按钮: ({btn_x}, {btn_y})")
        pmouse.double_click(coords=(btn_x, btn_y))
        time.sleep(1.5)

        hwnds = find_windows(title="新建工单")
        if hwnds:
            logger.info("新建工单已打开")
            return True

        logger.error("新建工单未出现")
        return False

    def find_work_order_window(self) -> Optional[int]:
        """
        查找已打开的工单窗口（标题包含 "GD" + "工作单"）。
        返回 hwnd，找不到返回 None。找到后自动最大化。
        """
        user32 = ctypes.windll.user32
        found = []

        def enum_callback(hwnd, lParam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buf, 256)
            title = buf.value
            if "GD" in title and "工作" in title:
                found.append(hwnd)
                return False
            return True

        CB = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(CB(enum_callback), 0)

        if not found:
            logger.error("未找到已打开的工单窗口")
            return None

        hwnd = found[0]
        buf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, buf, 256)
        logger.info(f"找到工单窗口: {buf.value} (hwnd={hwnd})")

        # 最大化
        user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
        time.sleep(0.5)
        logger.info("工单窗口已最大化")
        return hwnd

    def click_new_row(self, hwnd: int) -> bool:
        """
        在工单窗口 DataGridView 中点击 NewItem Row 的第一个单元格（编号列）。
        返回 True 表示定位成功。
        """
        import uiautomation as uia

        try:
            el = uia.ControlFromHandle(hwnd)
            grid = el.Control(AutomationId="gridControl")
            if not grid.Exists(0):
                logger.error("未找到 gridControl")
                return False

            data_panel = grid.Control(Name="Data Panel")
            if not data_panel.Exists(0):
                logger.error("未找到 Data Panel")
                return False

            rows = data_panel.GetChildren()
            logger.info(f"DataGridView 可见行数: {len(rows)}")

            # 找到 NewItem Row
            new_row = None
            for r in rows:
                name = r.Name.lower() if r.Name else ""
                if "new" in name or "newitem" in name:
                    new_row = r
                    break
            if new_row is None:
                new_row = rows[-1]

            logger.info(f"目标行: {new_row.Name}")

            # 点击第一个单元格（编号列）
            cells = new_row.GetChildren()
            if not cells:
                logger.error("NewItem Row 无单元格")
                return False

            first_cell = cells[0]
            cell_rect = first_cell.BoundingRectangle
            cx = cell_rect.left + cell_rect.width() // 2
            cy = cell_rect.top + cell_rect.height() // 2
            logger.info(f"点击 NewItem Row 编号列: ({cx}, {cy})")

            user32 = ctypes.windll.user32
            user32.SetCursorPos(cx, cy)
            time.sleep(0.1)
            # 单击
            user32.mouse_event(0x0002, 0, 0, 0, 0)  # left down
            time.sleep(0.05)
            user32.mouse_event(0x0004, 0, 0, 0, 0)  # left up
            time.sleep(0.2)
            logger.info("NewItem Row 定位完成")
            return True

        except Exception as e:
            logger.error(f"定位 NewItem Row 失败: {e}")
            return False

    # ------------------------------------------------------------------
    # 单行录入
    # ------------------------------------------------------------------

    def _enter_code(self, code: str):
        """粘贴产品代码并回车（代码查询需要较长延时）。"""
        self._paste_and_enter(str(code), pre_delay=0.15, post_delay=0.25)

    def _enter_value(self, value: str):
        """粘贴数值并回车（数量/份数/页数，延时较短）。"""
        self._paste_and_enter(str(value), pre_delay=0.03, post_delay=0.1)

    def _paste_and_enter(self, text: str, pre_delay: float = 0.03, post_delay: float = 0.1):
        """剪贴板粘贴 + 回车，可配置延时（补录模式叠加 _extra_delay）。"""
        pyperclip.copy(text)
        keyboard.send_keys("^v")
        time.sleep(pre_delay)
        keyboard.send_keys("{ENTER}")
        time.sleep(post_delay + self._extra_delay)

    def _fill_row(self, code: str, copies: str, pages: str):
        """填写一行数据(只填份数和页数，跳过价格)。"""
        self._enter_code(code)
        # 跳过价格(实价)，直接回车
        keyboard.send_keys("{ENTER}")
        time.sleep(0.1 + self._extra_delay)
        self._enter_value(copies)
        self._enter_value(pages)

    def _go_to_next_row(self):
        """从当前行末尾跳到下一行编号位置。"""
        keyboard.send_keys("{ENTER}")
        time.sleep(0.1 + self._extra_delay)

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------

    def run(self, records: List[Dict], append_mode: bool = False) -> Dict[str, int]:
        """
        执行批量网格录入。
        append_mode=True 时跳过第一行鼠标点击，从当前光标位置直接开始录入。
        """
        code_map = self.config.get("product_codes", {})
        columns = self.config.get("columns", {})
        success = 0
        failed = 0

        # 补录模式放慢：每一步叠加 append_delay
        self._extra_delay = self.append_delay if append_mode else 0.0

        size_col = columns.get("尺寸分类", "尺寸分类")
        subtype_col = columns.get("子项", "子项")
        qty_col = columns.get("数量", "数量")
        copies_col = columns.get("份数", "份数")
        pages_col = columns.get("页数", "页数")

        logger.info(f"{'=' * 50}")
        logger.info(f"开始批量录入，共 {len(records)} 条")
        logger.info(f"{'=' * 50}")

        for i, row in enumerate(records):
            logger.info(f"--- 第 {i + 1}/{len(records)} 条 ---")
            try:
                # 查找产品代码
                size = str(row.get(size_col, "")).strip()
                subtype = str(row.get(subtype_col, "")).strip() if subtype_col and subtype_col in row else ""
                if not subtype:
                    subtype = self.config.get("default_subtype", "")
                code_key = f"{subtype}_{size}" if subtype else size
                code = code_map.get(code_key, "")
                if not code and subtype:
                    # 回退1: 只按尺寸查找
                    code = code_map.get(size, "")
                if not code and subtype:
                    # 回退2: 只按子项查找 (如硫酸纸不分尺寸)
                    code = code_map.get(subtype, "")

                if not code:
                    raise ValueError(f"未找到产品代码: {code_key} (子项={subtype}, 尺寸={size})")

                copies = str(row.get(copies_col, "")).strip()
                pages = str(row.get(pages_col, "")).strip()

                logger.info(f"  {subtype}/{size} → {code} | 份数={copies} 页数={pages}")

                if i == 0:
                    if append_mode:
                        # 补录模式: 光标已在 NewItem Row 上，直接开始录入
                        logger.debug("  补录模式: 跳过首行点击，从当前光标位置开始")
                    else:
                        # 第一行: 点击编号
                        wx, wy = self.window.rectangle().left, self.window.rectangle().top
                        sx = wx + self.rel_x
                        sy = wy + self.rel_y_base
                        logger.debug(f"  点击: ({sx}, {sy})")
                        mouse.click(coords=(sx, sy))
                        time.sleep(0.3)
                else:
                    # 后续行: Enter 跳到下一行编号
                    self._go_to_next_row()

                self._fill_row(code, copies, pages)
                success += 1

            except Exception as e:
                failed += 1
                logger.error(f"  失败: {e}")
                keyboard.send_keys("{ESC}")
                time.sleep(0.3)

        # 最后一行完成后补一次回车
        if success > 0:
            keyboard.send_keys("{ENTER}")
            time.sleep(0.1)

        logger.info(f"{'=' * 50}")
        logger.info(f"完成: 成功 {success}, 失败 {failed}")
        logger.info(f"{'=' * 50}")
        return {"success": success, "failed": failed}
