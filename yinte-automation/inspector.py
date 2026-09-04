"""
UI 探查工具 — 列出系统中所有顶层窗口，并导出目标窗口的完整控件树。

使用方法:
    python inspector.py                           # 先列出所有窗口
    python inspector.py --title "印特"             # 指定窗口标题，导出控件树
    python inspector.py --title "印特" --output    # 同时保存到文件
"""

import argparse
import sys
import json
from pathlib import Path
from datetime import datetime

# 尝试用 uia 后端，失败则用 win32
BACKENDS = ["uia", "win32"]


def list_all_windows():
    """列出系统中所有顶层窗口。"""
    from pywinauto import Desktop

    print("\n" + "=" * 70)
    print("系统中所有顶层窗口:")
    print("=" * 70)

    windows = []
    for backend in BACKENDS:
        try:
            desktop = Desktop(backend=backend)
            for w in desktop.windows():
                try:
                    title = w.window_text()
                    class_name = w.class_name() if hasattr(w, 'class_name') else "?"
                    if title:
                        windows.append((title, class_name, backend))
                except Exception:
                    pass
        except Exception as e:
            print(f"  [注意] {backend} 后端不可用: {e}")

    # 去重 + 排序
    seen = set()
    for title, cls, backend in sorted(windows, key=lambda x: x[0]):
        key = title.strip()
        if key not in seen:
            seen.add(key)
            print(f"  标题: '{key}'")
            print(f"    类名: {cls}, 后端: {backend}")
            print()

    return list(seen)


def inspect_window(window_title: str, output_file: str = None):
    """导出目标窗口的控件树。"""
    from pywinauto import Application
    from pywinauto.findwindows import find_windows

    print(f"\n正在搜索窗口标题包含 '{window_title}' 的程序...")

    app = None
    window = None
    used_backend = None

    for backend in BACKENDS:
        try:
            app = Application(backend=backend).connect(
                title_re=f".*{window_title}.*"
            )
            window = app.window(title_re=f".*{window_title}.*")
            window.wait("visible", timeout=5)
            used_backend = backend
            break
        except Exception:
            continue

    if app is None:
        print(f"\n[错误] 未找到窗口标题包含 '{window_title}' 的程序。")
        print("请确认印特软件已打开，且窗口标题正确。")
        print("\n尝试列出所有窗口:")
        list_all_windows()
        return

    print(f"\n已连接 (后端: {used_backend})")
    print(f"窗口标题: '{window.window_text()}'")

    # 递归打印控件树
    print("\n" + "=" * 70)
    print("控件树:")
    print("=" * 70)

    tree_str = _dump_tree(window, indent=0)

    print(tree_str)

    # 如果指定输出文件
    if output_file:
        path = Path(output_file)
        path.write_text(
            f"# 印特 GUI 控件探查结果\n"
            f"# 时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"# 窗口标题: {window.window_text()}\n"
            f"# 后端: {used_backend}\n"
            f"{'=' * 70}\n"
            f"{tree_str}\n",
            encoding="utf-8",
        )
        print(f"\n控件树已保存到: {path.resolve()}")

    # 额外：打印一个简化列表（仅 Edit/ComboBox/Button）
    print("\n" + "=" * 70)
    print("简化控件列表 (仅 Edit / ComboBox / CheckBox / Button / DateTimePicker):")
    print("=" * 70)
    print(f"{'索引':<5} {'类型':<20} {'标题':<30} {'Auto ID':<30} {'类名'}")
    print("-" * 100)

    idx = 0
    interest_types = {"Edit", "ComboBox", "CheckBox", "Button", "DateTimePicker"}
    lines = []
    try:
        for c in window.descendants():
            try:
                ct = c.element_info.control_type
                if ct not in interest_types:
                    continue
                title = c.window_text()
                try:
                    auto_id = c.automation_id()
                except Exception:
                    auto_id = "N/A"
                cls = c.class_name() if hasattr(c, 'class_name') else "?"
                line = f"{idx:<5} {ct:<20} {title:<30} {auto_id:<30} {cls}"
                print(line)
                lines.append(line)
                idx += 1
            except Exception:
                pass
    except Exception as e:
        print(f"  遍历异常: {e}")

    if output_file:
        with open(output_file, "a", encoding="utf-8") as f:
            f.write(f"\n{'=' * 70}\n")
            f.write("简化控件列表:\n")
            f.write(f"{'索引':<5} {'类型':<20} {'标题':<30} {'Auto ID':<30} {'类名'}\n")
            f.write("-" * 100 + "\n")
            f.writelines(l + "\n" for l in lines)


def _dump_tree(element, indent: int = 0) -> str:
    """递归生成控件树的文本表示。"""
    prefix = "  " * indent
    lines = []
    try:
        title = element.window_text()
        ct = element.element_info.control_type
        try:
            auto_id = element.automation_id()
        except Exception:
            auto_id = "N/A"
        cls = element.class_name() if hasattr(element, 'class_name') else "?"

        node = f"{prefix}├─ [{ct}] title='{title}' auto_id='{auto_id}' class='{cls}'"
        lines.append(node)

        for child in element.children():
            lines.append(_dump_tree(child, indent + 1))
    except Exception:
        pass
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="印特 GUI 控件探查工具")
    parser.add_argument(
        "--title", "-t",
        type=str,
        help="窗口标题关键词（如 '印特'），不指定则列出所有窗口",
    )
    parser.add_argument(
        "--output", "-o",
        action="store_true",
        help="同时输出到 inspector_output.txt 文件",
    )
    args = parser.parse_args()

    if args.title:
        output_file = "inspector_output.txt" if args.output else None
        inspect_window(args.title, output_file)
    else:
        print("未指定 --title，列出所有窗口。")
        print("提示: python inspector.py --title '印特' 来探查目标窗口。")
        list_all_windows()


if __name__ == "__main__":
    main()
