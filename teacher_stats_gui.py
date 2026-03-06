#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教师科研统计分析工具 - 图形界面
"""

import PySimpleGUI as sg
import threading
import subprocess
import platform
import queue
import sys
import io
from pathlib import Path

try:
    import tomllib
except ImportError:
    import tomli as tomllib

import teacher_stats

DEFAULT_DATA_PATH = teacher_stats.DATA_PATH


def _get_documents_path():
    """获取用户文档文件夹路径（跨平台）"""
    if platform.system() == "Windows":
        try:
            import ctypes.wintypes
            CSIDL_PERSONAL = 5
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, 0, buf)
            return Path(buf.value)
        except Exception:
            return Path.home() / "Documents"
    return Path.home() / "Documents"


DEFAULT_OUTPUT_BASE = _get_documents_path()

# 数据文件：key -> (显示名称, 默认文件名, 文件类型)
DATA_FILE_INFO = {
    "在编信息汇总": ("在编信息汇总", "在编信息汇总.xlsx", (("Excel Files", "*.xlsx"), ("All Files", "*.*"))),
    "人才信息汇总": ("人才信息汇总", "人才信息汇总.xlsx", (("Excel Files", "*.xlsx"), ("All Files", "*.*"))),
    "成果批量导出": ("成果批量导出", "成果批量导出.xlsx", (("Excel Files", "*.xlsx"), ("All Files", "*.*"))),
    "纵向项目": ("纵向项目", "纵向项目.xls", (("Excel Files", "*.xls"), ("All Files", "*.*"))),
    "横向项目": ("横向项目", "横向项目.xls", (("Excel Files", "*.xls"), ("All Files", "*.*"))),
}


class QueueWriter(io.TextIOBase):
    """将 write 调用转发到 queue，供主线程实时读取"""
    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(text)
        return len(text) if text else 0

    def flush(self):
        pass


SUBDIR_NAME = "teacher_stats_output"
FONT_MAIN = ("Arial", 16)
FONT_LOG = ("Arial", 12)


def _get_app_path():
    """获取应用资源路径（兼容 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


CONFIG_PATH = _get_app_path() / "teacher_stats_config.toml"


def get_default_config():
    """返回默认配置（始终由代码生成，不依赖外部文件）"""
    cfg = {
        "teachers": [],
        "output_base": str(DEFAULT_OUTPUT_BASE),
        "use_subdir": True,
        "files": {},
    }
    # 仅当 data 目录存在时（本地开发）填充默认文件路径，否则留空让用户自行选择
    data_exists = DEFAULT_DATA_PATH.is_dir()
    for key, (_, default_name, _) in DATA_FILE_INFO.items():
        cfg["files"][key] = str(DEFAULT_DATA_PATH / default_name) if data_exists else ""
    return cfg


def _escape_toml_str(s):
    return str(s).replace("\\", "\\\\").replace('"', '\\"')


def _format_toml_array(items):
    if not items:
        return "[]"
    return "[" + ", ".join(f'"{_escape_toml_str(x)}"' for x in items) + "]"


def load_config():
    """加载上次退出时的缓存配置；若缓存不存在则返回代码中的默认配置"""
    default_cfg = get_default_config()
    if not CONFIG_PATH.exists():
        return default_cfg
    try:
        with open(CONFIG_PATH, "rb") as f:
            data = tomllib.load(f)
    except Exception:
        return default_cfg
    last = data.get("last")
    if not last:
        return default_cfg
    cached_output = last.get("output_base", "")
    # 缓存的路径可能来自其他操作系统，验证其有效性
    if cached_output and Path(cached_output).exists():
        output_base = cached_output
    else:
        output_base = default_cfg["output_base"]

    cfg = {
        "teachers": last.get("teachers", default_cfg["teachers"]),
        "output_base": output_base,
        "use_subdir": last.get("use_subdir", default_cfg["use_subdir"]),
        "files": dict(default_cfg["files"]),
    }
    for key in DATA_FILE_INFO:
        cached_file = last.get("files", {}).get(key, "")
        if cached_file and Path(cached_file).exists():
            cfg["files"][key] = cached_file
        else:
            cfg["files"][key] = default_cfg["files"][key]
    return cfg


def save_config(last_cfg):
    """将上次退出时的界面状态缓存到 TOML 文件（默认值始终由代码生成）"""
    lines = ["# 此文件由程序自动生成，仅缓存上次退出时的界面状态", ""]
    lines.append("[last]")
    lines.append(f"teachers = {_format_toml_array(last_cfg['teachers'])}")
    lines.append(f'output_base = "{_escape_toml_str(last_cfg["output_base"])}"')
    lines.append(f'use_subdir = {"true" if last_cfg["use_subdir"] else "false"}')
    lines.append("")
    lines.append("[last.files]")
    for key, path in last_cfg["files"].items():
        lines.append(f'"{_escape_toml_str(key)}" = "{_escape_toml_str(path)}"')
    lines.append("")
    CONFIG_PATH.write_text("\n".join(lines), encoding="utf-8")


def collect_config(values):
    """从当前窗口值中收集配置"""
    raw = values["-TEACHERS-"].strip()
    names = [n.strip() for n in raw.replace("，", ",").replace("\n", ",").split(",") if n.strip()]
    cfg = {
        "teachers": names,
        "output_base": values["-OUTPUT_BASE-"].strip(),
        "use_subdir": values["-USE_SUBDIR-"],
        "files": {},
    }
    for key in DATA_FILE_INFO:
        cfg["files"][key] = values[f"-FILE_{key}-"].strip()
    return cfg


def apply_config(window, config):
    """将配置应用到窗口"""
    window["-TEACHERS-"].update("\n".join(config.get("teachers", [])))
    window["-OUTPUT_BASE-"].update(config.get("output_base", ""))
    window["-USE_SUBDIR-"].update(config.get("use_subdir", True))
    for key in DATA_FILE_INFO:
        window[f"-FILE_{key}-"].update(config.get("files", {}).get(key, ""))
    # 更新输出目录预览
    base = config.get("output_base", "")
    use_subdir = config.get("use_subdir", True)
    if base and use_subdir:
        effective = str(Path(base) / SUBDIR_NAME)
        exists = (Path(base) / SUBDIR_NAME).is_dir()
        status = "✅ 已存在" if exists else "⚠️ 不存在，运行时自动创建"
    elif base:
        effective = base
        status = ""
    else:
        effective = ""
        status = ""
    window["-EFFECTIVE_PATH-"].update(effective)
    window["-SUBDIR_STATUS-"].update(status)
    for k in window.key_dict:
        elem = window[k]
        if isinstance(elem, sg.Input) and hasattr(elem, 'Widget') and elem.Widget:
            elem.Widget.xview_moveto(1.0)


def create_layout():
    default_base = str(DEFAULT_OUTPUT_BASE)
    default_effective = str(Path(default_base) / SUBDIR_NAME)
    base_has_subdir = (Path(default_base) / SUBDIR_NAME).is_dir()
    status_text = "✅ 已存在" if base_has_subdir else "⚠️ 不存在，运行时自动创建"

    # --- 左栏：数据文件、输出目录、复选框 ---
    file_rows = []
    for key, (label, default_name, file_types) in DATA_FILE_INFO.items():
        default_path = str(DEFAULT_DATA_PATH / default_name)
        file_rows.append([
            sg.Text(f"{label}:", size=(14, 1)),
            sg.Input(default_text=default_path, size=(36, 1), key=f"-FILE_{key}-"),
            sg.FileBrowse("浏览", file_types=file_types, target=f"-FILE_{key}-"),
            sg.Button("打开", key=f"-OPEN_FILE_DIR_{key}-"),
        ])
    file_frame = sg.Frame("数据文件", file_rows, expand_x=True)

    output_frame = sg.Frame("输出目录", [
        [sg.Text("基础目录:", size=(14, 1)),
         sg.Input(default_text=default_base, size=(36, 1), key="-OUTPUT_BASE-", enable_events=True),
         sg.FolderBrowse("浏览")],
        [sg.Checkbox(f"输出到 {SUBDIR_NAME} 子文件夹", default=True,
                     key="-USE_SUBDIR-", enable_events=True),
         sg.Text(status_text, key="-SUBDIR_STATUS-", size=(22, 1))],
        [sg.Text("实际输出路径:")],
        [sg.Input(default_text=default_effective, size=(58, 1),
                  key="-EFFECTIVE_PATH-", disabled=True,
                  text_color="red", disabled_readonly_background_color="#f0f0f0")],
    ], expand_x=True)

    left_col = sg.Column([
        [file_frame],
        [output_frame],
    ], vertical_alignment='top')

    # --- 右栏：教师姓名、运行日志 ---
    teacher_frame = sg.Frame("教师姓名", [
        [sg.Text("输入教师姓名（每行一个或用逗号分隔）：")],
        [sg.Multiline(default_text="", size=(40, 8), key="-TEACHERS-")],
    ], expand_x=True)

    log_frame = sg.Frame("运行日志", [
        [sg.Multiline(size=(40, 15), key="-LOG-", disabled=True, autoscroll=True, font=FONT_LOG)],
    ], expand_x=True)

    right_col = sg.Column([
        [teacher_frame],
        [log_frame],
    ], vertical_alignment='top')

    layout = [
        [left_col, right_col],
        [sg.Button("运行", key="-RUN-", size=(8, 1)),
         sg.Button("院系统计", key="-DEPT_STATS-", size=(10, 1)),
         sg.Button("刷新", key="-REFRESH-", size=(8, 1)),
         sg.Button("恢复默认配置", key="-RESET_DEFAULT-"),
         sg.Button("恢复上次配置", key="-RESET_LAST-"),
         sg.Button("打开输出目录", key="-OPEN_DIR-", size=(14, 1)),
         sg.Button("退出", key="-EXIT-", size=(8, 1))],
    ]
    return layout


def update_output_preview(window, values):
    """根据基础目录和复选框状态，更新实际输出路径和状态提示"""
    base = values["-OUTPUT_BASE-"].strip()
    use_subdir = values["-USE_SUBDIR-"]

    if not base:
        window["-EFFECTIVE_PATH-"].update("")
        window["-SUBDIR_STATUS-"].update("")
        return

    if use_subdir:
        effective = str(Path(base) / SUBDIR_NAME)
        exists = (Path(base) / SUBDIR_NAME).is_dir()
        status = "✅ 已存在" if exists else "⚠️ 不存在，运行时自动创建"
    else:
        effective = base
        status = ""

    window["-EFFECTIVE_PATH-"].update(effective)
    window["-SUBDIR_STATUS-"].update(status)


def open_folder(path):
    """用系统文件管理器打开文件夹"""
    system = platform.system()
    if system == "Darwin":
        subprocess.Popen(["open", path])
    elif system == "Windows":
        subprocess.Popen(["explorer", path])
    else:
        subprocess.Popen(["xdg-open", path])


def run_analysis(window, log_queue, teacher_names, file_paths, output_path):
    """在子线程中运行分析，stdout 实时写入 queue"""
    old_stdout = sys.stdout
    sys.stdout = QueueWriter(log_queue)
    try:
        result_dir = teacher_stats.main(
            teacher_names=teacher_names,
            file_paths=file_paths,
            output_path=output_path,
        )
        window.write_event_value("-DONE-", result_dir)
    except Exception as e:
        print(f"\n\n错误: {e}")
        window.write_event_value("-ERROR-", str(e))
    finally:
        sys.stdout = old_stdout


def run_department_analysis(window, log_queue, file_paths, output_path):
    """在子线程中运行院系统计，stdout 实时写入 queue"""
    old_stdout = sys.stdout
    sys.stdout = QueueWriter(log_queue)
    try:
        result_dir = teacher_stats.run_department_stats(
            file_paths=file_paths,
            output_path=output_path,
        )
        window.write_event_value("-DEPT_DONE-", result_dir)
    except Exception as e:
        print(f"\n\n错误: {e}")
        window.write_event_value("-DEPT_ERROR-", str(e))
    finally:
        sys.stdout = old_stdout


def main():
    sg.theme("LightGrey1")
    sg.set_options(font=FONT_MAIN)
    window = sg.Window("教师科研统计分析工具", create_layout(), finalize=True)

    # 加载并应用上次退出时的配置
    startup_config = load_config()
    apply_config(window, startup_config)

    running = False
    log_queue = queue.Queue()
    last_values = None

    while True:
        event, values = window.read(timeout=100)
        if values:
            last_values = values

        if event in (sg.WIN_CLOSED, "-EXIT-"):
            # 退出时保存当前配置
            try:
                v = last_values or values
                if v and "-TEACHERS-" in v:
                    save_config(collect_config(v))
            except Exception:
                pass
            break

        # 实时刷新日志
        updated = False
        while not log_queue.empty():
            try:
                text = log_queue.get_nowait()
                window["-LOG-"].update(text, append=True)
                updated = True
            except queue.Empty:
                break

        # 基础目录或复选框变化时更新预览
        if event in ("-OUTPUT_BASE-", "-USE_SUBDIR-"):
            update_output_preview(window, values)

        if event == "-RUN-" and not running:
            raw = values["-TEACHERS-"].strip()
            names = [n.strip() for n in raw.replace("，", ",").replace("\n", ",").split(",") if n.strip()]

            if not names:
                sg.popup_error("请输入至少一个教师姓名！")
                continue

            file_paths = {}
            missing = []
            for key in DATA_FILE_INFO:
                p = values[f"-FILE_{key}-"].strip()
                if not p or not Path(p).exists():
                    missing.append(DATA_FILE_INFO[key][0])
                file_paths[key] = p

            if missing:
                sg.popup_error(f"以下文件不存在：\n" + "\n".join(missing))
                continue

            effective_path = values["-EFFECTIVE_PATH-"].strip()
            if not effective_path:
                sg.popup_error("请选择输出目录！")
                continue

            # 检查教师是否在在编信息中
            try:
                not_found = teacher_stats.validate_teacher_names(names, file_paths)
            except Exception as e:
                sg.popup_error(f"校验教师姓名时出错：{e}")
                continue
            if not_found:
                answer = sg.popup_yes_no(
                    f"以下教师不在在编信息中：\n\n{'、'.join(not_found)}\n\n是否仍然继续？",
                    title="教师姓名校验",
                )
                if answer != "Yes":
                    continue
                log_queue.put(f"⚠️ 警告：以下教师不在在编信息中：{'、'.join(not_found)}\n")

            running = True
            window["-RUN-"].update(disabled=True)
            window["-DEPT_STATS-"].update(disabled=True)
            window["-LOG-"].update("")
            log_queue.put(f"开始分析：{', '.join(names)}\n")
            log_queue.put(f"输出目录：{effective_path}\n")

            thread = threading.Thread(
                target=run_analysis,
                args=(window, log_queue, names, file_paths, effective_path),
                daemon=True,
            )
            thread.start()

        if event == "-DEPT_STATS-" and not running:
            file_paths = {}
            missing = []
            for key in DATA_FILE_INFO:
                p = values[f"-FILE_{key}-"].strip()
                if not p or not Path(p).exists():
                    missing.append(DATA_FILE_INFO[key][0])
                file_paths[key] = p

            if missing:
                sg.popup_error(f"以下文件不存在：\n" + "\n".join(missing))
                continue

            effective_path = values["-EFFECTIVE_PATH-"].strip()
            if not effective_path:
                sg.popup_error("请选择输出目录！")
                continue

            running = True
            window["-RUN-"].update(disabled=True)
            window["-DEPT_STATS-"].update(disabled=True)
            window["-LOG-"].update("")
            log_queue.put("开始院系整体统计分析...\n")
            log_queue.put(f"输出目录：{effective_path}\n")

            thread = threading.Thread(
                target=run_department_analysis,
                args=(window, log_queue, file_paths, effective_path),
                daemon=True,
            )
            thread.start()

        if event == "-DONE-":
            result_dir = values[event]
            window["-LOG-"].update("\n✅ 分析完成！", append=True)
            running = False
            window["-RUN-"].update(disabled=False)
            window["-DEPT_STATS-"].update(disabled=False)
            # 刷新状态（文件夹已创建）
            update_output_preview(window, values)
            if result_dir:
                open_folder(result_dir)

        if event == "-ERROR-":
            window["-LOG-"].update("\n❌ 分析出错！", append=True)
            running = False
            window["-RUN-"].update(disabled=False)
            window["-DEPT_STATS-"].update(disabled=False)

        if event == "-DEPT_DONE-":
            result_dir = values[event]
            window["-LOG-"].update("\n✅ 院系统计完成！", append=True)
            running = False
            window["-RUN-"].update(disabled=False)
            window["-DEPT_STATS-"].update(disabled=False)
            update_output_preview(window, values)
            if result_dir:
                open_folder(result_dir)

        if event == "-DEPT_ERROR-":
            window["-LOG-"].update("\n❌ 院系统计出错！", append=True)
            running = False
            window["-RUN-"].update(disabled=False)
            window["-DEPT_STATS-"].update(disabled=False)

        if event == "-OPEN_DIR-":
            effective_path = values["-EFFECTIVE_PATH-"].strip()
            if effective_path and Path(effective_path).is_dir():
                open_folder(effective_path)
            else:
                sg.popup_error(f"输出目录不存在，请先运行程序。")

        # 打开数据文件所在文件夹
        if event.startswith("-OPEN_FILE_DIR_"):
            for key in DATA_FILE_INFO:
                if event == f"-OPEN_FILE_DIR_{key}-":
                    file_path = values[f"-FILE_{key}-"].strip()
                    if file_path:
                        folder = str(Path(file_path).parent)
                        if Path(folder).is_dir():
                            open_folder(folder)
                        else:
                            sg.popup_error(f"文件夹不存在：{folder}")
                    break

        # 恢复默认配置
        if event == "-RESET_DEFAULT-":
            apply_config(window, get_default_config())
            window["-LOG-"].update("🔄 已恢复默认配置\n", append=True)

        # 恢复上次退出时配置
        if event == "-RESET_LAST-":
            apply_config(window, load_config())
            window["-LOG-"].update("🔄 已恢复上次退出时配置\n", append=True)

        # 全局刷新
        if event == "-REFRESH-":
            update_output_preview(window, values)
            log_lines = ["🔄 刷新状态：\n"]
            for key, (label, _, _) in DATA_FILE_INFO.items():
                p = values[f"-FILE_{key}-"].strip()
                exists = Path(p).exists() if p else False
                status = "✅" if exists else "❌ 不存在"
                log_lines.append(f"  {label}: {status}\n")
            effective = values["-EFFECTIVE_PATH-"].strip()
            if effective:
                out_exists = Path(effective).is_dir()
                out_status = "✅" if out_exists else "⚠️ 不存在"
                log_lines.append(f"  输出目录: {out_status}\n")
            window["-LOG-"].update("".join(log_lines), append=True)
            for k in window.key_dict:
                elem = window[k]
                if isinstance(elem, sg.Input) and hasattr(elem, 'Widget') and elem.Widget:
                    elem.Widget.xview_moveto(1.0)

    window.close()


if __name__ == "__main__":
    main()
