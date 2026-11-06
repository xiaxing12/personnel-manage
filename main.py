# -*- coding: utf-8 -*-
"""员工数据管理系统：查询 / 结果 / 添加 三页桌面小程序。"""

import base64
import datetime
import io
import json
import os
import re
import sys
import tempfile
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from cryptography.fernet import Fernet, InvalidToken
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
except ImportError:  # pragma: no cover
    Fernet = InvalidToken = None
    hashes = PBKDF2HMAC = None

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    Workbook = load_workbook = None
    LineChart = Reference = ColorScaleRule = None
    DataValidation = None
    Font = PatternFill = Alignment = get_column_letter = None

try:
    from tkcalendar import DateEntry
except ImportError:  # pragma: no cover
    DateEntry = None

DATA_FILE_NAME = "员工数据.xlsx"
DEFAULT_COLUMNS = [
    "员工姓名",
    "日期",
    "工号",
    "是否新人",
    "是否支援人员",
    "班组名称",
    "产品客户",
    "所属机种",
    "物料代码",
    "所属产线/工序",
    "模具号",
    "当日计划产量",
    "实际产量",
    "合格品数量",
    "不良品数量",
    "达成率",
    "合格率",
    "工时",
    "备注",
]
REQUIRED_NAME = "员工姓名"
REQUIRED_DATE = "日期"
RATE_COLUMNS = ("达成率", "合格率")
YES_NO_COLUMNS = ("是否新人", "是否支援人员")
IMPORT_COLUMNS = tuple(
    column for column in DEFAULT_COLUMNS if column not in RATE_COLUMNS
)
IMPORT_HEADER_ALIASES = {
    "姓名": "员工姓名",
    "员工姓名": "员工姓名",
    "生产日期": "日期",
    "当日日期": "日期",
    "日期": "日期",
    "员工编号": "工号",
    "员工工号": "工号",
    "工号": "工号",
    "是否新员工": "是否新人",
    "是否新人": "是否新人",
    "是否支援": "是否支援人员",
    "是否支援人员": "是否支援人员",
    "组别": "班组名称",
    "班组": "班组名称",
    "班组名称": "班组名称",
    "客户": "产品客户",
    "产品客户": "产品客户",
    "机种": "所属机种",
    "所属机种": "所属机种",
    "物料": "物料代码",
    "物料代码": "物料代码",
    "工序": "所属产线/工序",
    "产线/工序": "所属产线/工序",
    "所属产线/工序": "所属产线/工序",
    "模具": "模具号",
    "模具号": "模具号",
    "计划产量": "当日计划产量",
    "当日计划产量": "当日计划产量",
    "实际产量": "实际产量",
    "合格数量": "合格品数量",
    "合格品数量": "合格品数量",
    "不良数量": "不良品数量",
    "不良品数量": "不良品数量",
    "工作时长": "工时",
    "工时": "工时",
    "说明": "备注",
    "备注": "备注",
}
RESULT_COLUMN_WIDTHS = {
    "员工姓名": 110,
    "日期": 110,
    "工号": 100,
    "是否新人": 90,
    "是否支援人员": 110,
    "班组名称": 120,
    "产品客户": 110,
    "所属机种": 120,
    "物料代码": 130,
    "所属产线/工序": 170,
    "模具号": 110,
    "当日计划产量": 145,
    "实际产量": 120,
    "合格品数量": 125,
    "不良品数量": 125,
    "达成率": 100,
    "合格率": 100,
    "工时": 90,
    "备注": 180,
}
RATE_FILL_ROWS = 2000
PROGRAM_KEY = os.environ.get("EMP_DATA_KEY", "DEMO-ONLY-DATA-KEY-CHANGE-ME")
RENEWAL_KEY = os.environ.get("EMP_RENEWAL_KEY", "DEMO-ONLY-RENEWAL-KEY-CHANGE-ME")
LICENSE_EXPIRE_DATE = datetime.date.fromisoformat(os.environ.get("EMP_LICENSE_EXPIRE_DATE", "2099-12-31"))
MAX_ATTEMPTS = 3
LOCK_HOURS = 24
STATE_FILE = "程序授权.dat"
APP_TITLE = "新人员管理系统"
FONT_FAMILY = "Microsoft YaHei UI"
OPTION_FILES = {
    "物料代码": "物料代码选项.txt",
    "员工姓名": "员工姓名选项.txt",
    "工号": "工号选项.txt",
    "班组名称": "班组名称选项.txt",
    "所属产线/工序": "所属产线工序选项.txt",
    "模具号": "模具号选项.txt",
    "产品客户": "产品客户选项.txt",
    "所属机种": "所属机种选项.txt",
}
OPTION_FIELDS = (
    "员工姓名",
    "工号",
    "班组名称",
    "所属机种",
    "物料代码",
    "产品客户",
    "所属产线/工序",
    "模具号",
)
OPTION_STORE_FILE = "选项内容.dat"
TEMPLATE_DIR = "表格模板"


class AppError(Exception):
    """业务错误，用于向用户展示友好提示。"""


MAGIC = b"EMPDATA1\x00"


def derive_key(secret, salt):
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(), length=32, salt=salt, iterations=200000
    )
    return base64.urlsafe_b64encode(kdf.derive(secret.encode("utf-8")))


def encrypt_bytes(secret, data):
    """用密钥加密字节数据，返回带标识与盐的密文。"""
    salt = os.urandom(16)
    key = derive_key(secret, salt)
    token = Fernet(key).encrypt(data)
    return MAGIC + salt + token


def decrypt_bytes(secret, data):
    """用密钥解密；密钥错误时抛出 AppError。"""
    if not data.startswith(MAGIC):
        raise AppError("数据文件不是加密格式。")
    salt = data[len(MAGIC):len(MAGIC) + 16]
    token = data[len(MAGIC) + 16:]
    key = derive_key(secret, salt)
    try:
        return Fernet(key).decrypt(token)
    except InvalidToken:
        raise AppError("密钥错误，无法打开数据文件。")


def enable_dpi_awareness():
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def normalize_date(value):
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for sep in ("-", "/", "."):
        parts = text.split(sep)
        if len(parts) == 3:
            try:
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                return datetime.date(y, m, d).isoformat()
            except ValueError:
                continue
    return None


def normalize_import_date(value):
    """Excel 导入只接受日期单元格或严格的 YYYY-MM-DD 文本。"""
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if value is None:
        return None
    text = str(value).strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return None
    try:
        return datetime.date.fromisoformat(text).isoformat()
    except ValueError:
        return None


def to_num(value):
    if value is None:
        return None
    try:
        text = str(value).strip().replace(",", "")
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def calc_rates(plan, actual, good, bad):
    """根据产量/合格/不良数量计算达成率与合格率，返回显示文本。"""
    plan = to_num(plan)
    actual = to_num(actual)
    good = to_num(good)
    bad = to_num(bad)

    rate = "—"
    if plan is not None and actual is not None and plan > 0:
        rate = f"{actual / plan * 100:.1f}%"

    quality = "—"
    denom = actual if actual is not None and actual > 0 else None
    if denom is None and good is not None and bad is not None and good + bad > 0:
        denom = good + bad
    if denom is not None and good is not None and denom > 0:
        quality = f"{good / denom * 100:.1f}%"
    return rate, quality


def rate_formulas(row_idx, columns):
    """根据表头生成“达成率”“合格率”的 Excel 公式，供表格自动计算。"""
    def col(name):
        return get_column_letter(columns.index(name) + 1) if name in columns else None

    plan = col("当日计划产量")
    actual = col("实际产量")
    good = col("合格品数量")
    bad = col("不良品数量")
    rate = ""
    quality = ""
    if plan and actual:
        rate = (
            f'=IF(AND(ISNUMBER({plan}{row_idx}),ISNUMBER({actual}{row_idx}),'
            f'{plan}{row_idx}>0),TEXT({actual}{row_idx}/{plan}{row_idx},"0.0%"),"")'
        )
    if actual and good and bad:
        quality = (
            f'=IF(AND(ISNUMBER({good}{row_idx}),ISNUMBER({bad}{row_idx}),'
            f'{good}{row_idx}+{bad}{row_idx}>0),'
            f'TEXT({good}{row_idx}/IF(AND(ISNUMBER({actual}{row_idx}),{actual}{row_idx}>0),'
            f'{actual}{row_idx},{good}{row_idx}+{bad}{row_idx}),"0.0%"),"")'
        )
    return rate, quality


def summarize_rows(rows, columns):
    """对查询结果做合计，返回 (合计行各列的值, 汇总数字文本)。"""
    def total_of(key):
        total = 0.0
        for r in rows:
            v = to_num(r.get(key))
            if v is not None:
                total += v
        return total

    def fmt_num(value):
        if value == int(value):
            return str(int(value))
        return f"{value:.1f}"

    plan_total = total_of("当日计划产量")
    actual_total = total_of("实际产量")
    good_total = total_of("合格品数量")
    bad_total = total_of("不良品数量")
    hours_total = total_of("工时")
    rate_text, quality_text = calc_rates(plan_total, actual_total, good_total, bad_total)

    total_values = []
    for col_name in columns:
        if col_name == "当日计划产量":
            total_values.append(fmt_num(plan_total))
        elif col_name == "实际产量":
            total_values.append(fmt_num(actual_total))
        elif col_name == "合格品数量":
            total_values.append(fmt_num(good_total))
        elif col_name == "不良品数量":
            total_values.append(fmt_num(bad_total))
        elif col_name == "达成率":
            total_values.append(rate_text)
        elif col_name == "合格率":
            total_values.append(quality_text)
        elif col_name == "工时":
            total_values.append(fmt_num(hours_total))
        else:
            total_values.append("")
    totals = (
        fmt_num(plan_total),
        fmt_num(actual_total),
        fmt_num(good_total),
        fmt_num(bad_total),
        fmt_num(hours_total),
        rate_text,
        quality_text,
    )
    return total_values, totals


def build_summary_text(totals):
    plan, actual, good, bad, hours, rate, quality = totals
    return (
        "汇总（合计）：计划产量 {plan}　实际产量 {actual}　合格品 {good}　不良品 {bad}　工时 {hours} 小时\n"
        "达成率 {rate}　合格率 {quality}（按合计计算）".format(
            plan=plan, actual=actual, good=good, bad=bad, hours=hours, rate=rate, quality=quality
        )
    )


def style_header_row(ws, count):
    if Font is None:
        return
    fill = PatternFill("solid", fgColor="2F6FED")
    for cell in ws[1][:count]:
        cell.font = Font(bold=True, color="FFFFFF", name="微软雅黑", size=11)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def _walk_widgets(widget):
    yield widget
    for child in widget.winfo_children():
        yield from _walk_widgets(child)


class FilterCombobox(ttk.Combobox):
    """支持输入关键字筛选，同时保留原有的只读下拉字段限制。"""

    _NAVIGATION_KEYS = {
        "Left",
        "Right",
        "Up",
        "Down",
        "Home",
        "End",
        "Prior",
        "Next",
        "Return",
        "Tab",
        "Escape",
        "Shift_L",
        "Shift_R",
        "Control_L",
        "Control_R",
        "Alt_L",
        "Alt_R",
    }

    def __init__(self, master=None, **kwargs):
        self._input_locked = not bool(kwargs.pop("allow_input", True))
        initial_state = str(kwargs.get("state", "normal"))
        self._allow_custom = initial_state != "readonly"
        self._all_values = tuple(str(value) for value in kwargs.get("values", ()))
        self._last_valid = ""
        self._typing = False
        self._keep_filter = False
        self._filter_job = None
        self._reopen_job = None
        self._popup_allowed = True
        self._pending_text = ""
        self._original_postcommand = kwargs.pop("postcommand", None)
        kwargs["state"] = "normal"
        kwargs["postcommand"] = self._before_post
        super().__init__(master, **kwargs)
        self._install_popdown_focus_bindings()
        self._last_valid = self.get().strip()
        self.bind("<KeyPress>", self._on_key_press, add="+")
        self.bind("<KeyRelease>", self._on_key_release, add="+")
        self.bind("<FocusIn>", self._on_focus_in, add="+")
        self.bind("<FocusOut>", self._on_focus_out, add="+")
        self.bind("<Button-1>", self._on_mouse_click, add="+")
        self.bind("<Escape>", self._on_escape, add="+")
        self.bind("<<ComboboxSelected>>", self._on_selected, add="+")

    def _install_popdown_focus_bindings(self):
        """让下拉列表显示时仍由输入框接收键盘输入。"""
        self.tk.call(
            "bind",
            "ComboboxListbox",
            "<Map>",
            "focus -force [::ttk::combobox::LBMaster %W]",
        )
        # Windows 默认在列表失焦时关闭下拉，会阻止切回输入框。
        self.tk.call("bind", "ComboboxListbox", "<FocusOut>", "")
        try:
            popdown = self.tk.call(
                "ttk::combobox::PopdownWindow", self._w
            )
            self.tk.call(
                "bind",
                popdown,
                "<Map>",
                "after idle [list wm attributes %W -topmost 0]",
            )
            click_command = self.register(self._on_popdown_click)
            self.tk.call(
                "bind",
                popdown,
                "<Button-1>",
                f"{click_command} %X %Y",
            )
        except tk.TclError:
            pass

    def set_options(self, values):
        self._all_values = tuple(str(value) for value in values)
        self["values"] = self._all_values

    def clear_selection(self):
        self._cancel_pending_filter()
        self._pending_text = ""
        self._last_valid = ""
        self._typing = False
        self.set("")
        self["values"] = self._all_values

    def _cancel_pending_filter(self):
        if self._filter_job is not None:
            try:
                self.after_cancel(self._filter_job)
            except tk.TclError:
                pass
            self._filter_job = None

    def _cancel_pending_reopen(self):
        if self._reopen_job is not None:
            try:
                self.after_cancel(self._reopen_job)
            except tk.TclError:
                pass
            self._reopen_job = None

    def _close_dropdown(self):
        try:
            self.tk.call("ttk::combobox::Unpost", self._w)
        except tk.TclError:
            pass

    def _before_post(self):
        if not self._keep_filter:
            self["values"] = self._all_values
        if self._original_postcommand:
            self._original_postcommand()

    def _on_key_press(self, event):
        if self._input_locked and (
            event.char or event.keysym in ("BackSpace", "Delete")
        ):
            return "break"
        if event.char or event.keysym in ("BackSpace", "Delete"):
            self._cancel_pending_filter()
            self._close_dropdown()
            self._pending_text = ""
            self._typing = True

    def _on_key_release(self, event):
        if not self._typing or event.keysym in self._NAVIGATION_KEYS:
            return
        self._cancel_pending_filter()
        if not self.get().strip():
            self._last_valid = ""
            self._pending_text = ""
        self._filter_job = self.after_idle(self._apply_filter)

    def _apply_filter(self):
        self._filter_job = None
        if not self._typing or not self._popup_allowed:
            return
        needle = self.get().strip().casefold()
        if not needle:
            self._typing = False
            self._close_dropdown()
            self["values"] = self._all_values
            return
        matches = tuple(
            value
            for value in self._all_values
            if needle in value.casefold()
        )
        self._close_dropdown()
        self["values"] = matches
        if matches:
            self._keep_filter = True
            try:
                self.tk.call("ttk::combobox::Post", self._w)
            finally:
                self._keep_filter = False
        self._typing = False

    def _on_focus_in(self, event):
        current = self.get().strip()
        if self._pending_text and (not current or current == self._last_valid):
            self.set(self._pending_text)
            current = self._pending_text
        if not current or current in self._all_values:
            self._last_valid = current
        self._typing = False

    def _on_mouse_click(self, event):
        self._cancel_pending_reopen()
        self._reopen_job = self.after_idle(self._reopen_after_click)

    def _reopen_after_click(self):
        self._reopen_job = None
        if not self._popup_allowed:
            return
        if not self.get().strip():
            return
        try:
            popdown = self.tk.call("ttk::combobox::PopdownWindow", self._w)
            if int(self.tk.call("winfo", "ismapped", popdown)):
                return
        except (tk.TclError, ValueError):
            pass
        if self._input_locked:
            self._keep_filter = False
            self["values"] = self._all_values
            self.tk.call("ttk::combobox::Post", self._w)
            return
        self._typing = True
        self._apply_filter()

    def _on_popdown_click(self, x_root, y_root):
        try:
            x = int(x_root)
            y = int(y_root)
            combo_x = int(self.tk.call("winfo", "rootx", self))
            combo_y = int(self.tk.call("winfo", "rooty", self))
            combo_w = int(self.tk.call("winfo", "width", self))
            combo_h = int(self.tk.call("winfo", "height", self))
            if (
                combo_x + combo_w - 48 <= x < combo_x + combo_w
                and combo_y <= y < combo_y + combo_h
            ):
                self.after_idle(self._open_all_options)
        except (tk.TclError, TypeError, ValueError):
            pass

    def _open_all_options(self):
        if not self._popup_allowed:
            return
        self._cancel_pending_filter()
        self._cancel_pending_reopen()
        self._typing = False
        self._close_dropdown()
        self["values"] = self._all_values
        self._keep_filter = True
        try:
            self.tk.call("ttk::combobox::Post", self._w)
        finally:
            self._keep_filter = False

    def _on_focus_out(self, event):
        self._cancel_pending_filter()
        self.after_idle(self._finish_focus_out)

    def _finish_focus_out(self):
        try:
            popdown = self.tk.call("ttk::combobox::PopdownWindow", self._w)
            if int(self.tk.call("winfo", "ismapped", popdown)):
                return
        except (tk.TclError, ValueError):
            pass
        value = self.get().strip()
        if not value:
            self._pending_text = ""
            self._last_valid = ""
            self.set("")
        elif value in self._all_values:
            self._pending_text = ""
            self._last_valid = value
        elif self._allow_custom:
            self._pending_text = ""
            self._last_valid = value
        else:
            self._pending_text = value
            self.set(self._last_valid)
        self._typing = False
        self["values"] = self._all_values

    def _on_escape(self, event):
        self._cancel_pending_filter()
        self._pending_text = ""
        self.set(self._last_valid)
        self._typing = False
        self["values"] = self._all_values

    def _on_selected(self, event):
        self._cancel_pending_filter()
        self._pending_text = ""
        self._typing = False
        self._last_valid = self.get().strip()
        self["values"] = self._all_values


def set_combo_values(combo, values):
    if isinstance(combo, FilterCombobox):
        combo.set_options(values)
    else:
        combo["values"] = tuple(values)


def _bind_clipboard(widget):
    for seq, event in (
        ("<Control-c>", "<<Copy>>"),
        ("<Control-v>", "<<Paste>>"),
        ("<Control-x>", "<<Cut>>"),
        ("<Control-a>", "<<SelectAll>>"),
    ):
        def handler(e, ev=event):
            e.widget.event_generate(ev)
            return "break"

        widget.bind(seq, handler)


def bind_clipboard_shortcuts(root):
    """给所有输入框/下拉框加上 Ctrl+C/V/X/A 复制粘贴快捷键。"""
    for widget in _walk_widgets(root):
        if widget.winfo_class() in ("TEntry", "TCombobox"):
            _bind_clipboard(widget)


def _block_wheel(widget):
    for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        widget.bind(seq, lambda e: "break")


def block_combobox_wheel(combo):
    """滚轮在选项框上不生效，只用于滚动界面。"""
    _block_wheel(combo)
    for child in combo.winfo_children():
        if child.winfo_class() == "TEntry":
            _block_wheel(child)
    try:
        pop = combo.tk.call("ttk::combobox::PopdownWindow", combo._w)
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            combo.tk.call("bind", pop, seq, "break")
    except tk.TclError:
        pass


def bind_combobox_behavior(root):
    for widget in _walk_widgets(root):
        cls = widget.winfo_class()
        if cls == "TCombobox":
            block_combobox_wheel(widget)
            if isinstance(widget, FilterCombobox) and not getattr(
                widget, "_no_clear", False
            ):
                widget.bind(
                    "<Delete>",
                    lambda e: (e.widget.clear_selection(), "break")[1],
                )
        elif cls == "TEntry" and isinstance(widget, DateEntry):
            for seq in ("<Delete>", "<BackSpace>"):
                widget.bind(seq, lambda e: (e.widget.delete(0, "end"), "break")[1])


def close_combobox_popdowns(root):
    for widget in _walk_widgets(root):
        if widget.winfo_class() != "TCombobox":
            continue
        if isinstance(widget, FilterCombobox):
            widget._cancel_pending_filter()
            widget._cancel_pending_reopen()
        try:
            widget.tk.call("ttk::combobox::Unpost", widget._w)
        except tk.TclError:
            pass
    try:
        root_path = str(root)
        pending = [root_path]
        while pending:
            path = pending.pop()
            for child in root.tk.splitlist(
                root.tk.call("winfo", "children", path)
            ):
                pending.append(child)
                try:
                    if root.tk.call("winfo", "class", child) != "ComboboxPopdown":
                        continue
                    root.tk.call("grab", "release", child)
                    root.tk.call("wm", "withdraw", child)
                    root.tk.call("wm", "attributes", child, "-topmost", 0)
                except tk.TclError:
                    continue
    except tk.TclError:
        pass


def get_date_or_none(entry):
    """读取日期框的日期；为空或无效时返回 None。"""
    try:
        return entry.get_date()
    except Exception:
        return None


class ToggleSwitch(tk.Canvas):
    """自定义开关旋钮：圆角滑块。"""

    def __init__(self, master, variable, command=None, width=54, height=28, bg="#ffffff"):
        super().__init__(
            master,
            width=width,
            height=height,
            bg=bg,
            highlightthickness=0,
            bd=0,
        )
        self.var = variable
        self.command = command
        self.w = width
        self.h = height
        self.pill_w = 46
        self.pill_h = 24
        self.pill_x = 2
        self.pill_y = (height - self.pill_h) // 2
        self.bind("<Button-1>", self._click)
        variable.trace_add("write", lambda *a: self._draw())
        self._draw()

    def _click(self, event):
        self.var.set(not self.var.get())
        if self.command:
            self.command()

    def _capsule(self, x1, y1, x2, y2, **kw):
        """绘制圆润的胶囊形按钮。"""
        r = min((x2 - x1) / 2, (y2 - y1) / 2)
        self.create_rectangle(x1 + r, y1, x2 - r, y2, **kw)
        self.create_oval(x1, y1, x1 + 2 * r, y1 + 2 * r, **kw)
        self.create_oval(x2 - 2 * r, y1, x2, y1 + 2 * r, **kw)
        self.create_oval(x1, y2 - 2 * r, x1 + 2 * r, y2, **kw)
        self.create_oval(x2 - 2 * r, y2 - 2 * r, x2, y2, **kw)

    def _draw(self):
        self.delete("all")
        on = bool(self.var.get())
        pill_color = "#22c55e" if on else "#d1d5db"
        self._capsule(
            self.pill_x,
            self.pill_y,
            self.pill_x + self.pill_w,
            self.pill_y + self.pill_h,
            fill=pill_color,
            outline=pill_color,
        )
        knob_r = 9
        knob_y = self.h / 2
        knob_x = (
            self.pill_x + self.pill_w - knob_r - 3
            if on
            else self.pill_x + knob_r + 3
        )
        self.create_oval(
            knob_x - knob_r,
            knob_y - knob_r,
            knob_x + knob_r,
            knob_y + knob_r,
            fill="#ffffff",
            outline="#ffffff",
        )


class RoundStatusToggle(tk.Canvas):
    """圆形状态灯：红色关闭，绿色启用。"""

    def __init__(self, master, variable, command=None, bg="#f9fdfc"):
        super().__init__(
            master,
            width=32,
            height=32,
            bg=bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.var = variable
        self.command = command
        self.bind("<Button-1>", self._click)
        variable.trace_add("write", lambda *args: self._draw())
        self._draw()

    def _draw(self):
        self.delete("all")
        enabled = bool(self.var.get())
        color = "#3eb56b" if enabled else "#dc4c4c"
        self.create_oval(
            5,
            5,
            27,
            27,
            fill=color,
            outline="#ffffff",
            width=2,
        )

    def _click(self, event):
        self.var.set(not self.var.get())
        if self.command:
            self.command()


class ClearableDateEntry(DateEntry):
    """支持清空的日期框：按 Delete 清空后不会在失焦时自动恢复。"""

    def _validate_date(self):
        if not self.get().strip():
            return True
        return super()._validate_date()

    def get_date(self):
        if not self.get().strip():
            return None
        return super().get_date()

    def drop_down(self):
        """打开日历；日期为空时以今天为初始选中，避免解析空文本报错。"""
        if self._calendar.winfo_ismapped():
            self._top_cal.withdraw()
            return
        self._validate_date()
        if self.get().strip():
            date = self.parse_date(self.get())
        else:
            date = self._date or datetime.date.today()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        if self.winfo_toplevel().attributes("-topmost"):
            self._top_cal.attributes("-topmost", True)
        else:
            self._top_cal.attributes("-topmost", False)
        self._top_cal.geometry("+%i+%i" % (x, y))
        self._top_cal.deiconify()
        self._calendar.focus_set()
        self._calendar.selection_set(date)


def load_option_values(header):
    """从程序同目录的文本文件读取下拉备选值（每行一个）。"""
    filename = OPTION_FILES.get(header)
    if not filename:
        return []
    path = os.path.join(base_dir(), filename)
    values = []
    seen = set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                text = line.strip()
                if text and text not in seen:
                    seen.add(text)
                    values.append(text)
    except OSError:
        pass
    return values


def merge_options(header, existing):
    """把数据表里的值与备选文件里的值合并（保留顺序、去重）。"""
    result = list(existing)
    seen = set(result)
    for value in load_option_values(header):
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


class DataManager:
    """负责读写 Excel 数据文件。"""

    def __init__(self, path, secret):
        self.path = path
        self.secret = secret
        self._cache_signature = None
        self._cache_headers = ()
        self._cache_rows = []
        self._cache_column_values = {}
        self.ensure_file()

    def _workbook(self):
        return load_workbook(io.BytesIO(self._decrypted_bytes()))

    def _file_signature(self):
        """用于发现数据表被程序外部的解密工具更新。"""
        try:
            stat = os.stat(self.path)
            return stat.st_mtime_ns, stat.st_size
        except OSError:
            return None

    def file_signature(self):
        return self._file_signature()

    def _invalidate_cache(self):
        self._cache_signature = None
        self._cache_headers = ()
        self._cache_rows = []
        self._cache_column_values = {}

    def _decrypted_bytes(self):
        with open(self.path, "rb") as f:
            raw = f.read()
        if raw.startswith(MAGIC):
            return decrypt_bytes(self.secret, raw)
        return raw

    def _headers_from_sheet(self, ws):
        headers = []
        for cell in ws[1]:
            headers.append(str(cell.value).strip() if cell.value is not None else "")
        return headers or list(DEFAULT_COLUMNS)

    def _snapshot(self):
        """一次解密并解析数据表；文件未变化时直接复用解析结果。"""
        signature = self._file_signature()
        if signature is not None and signature == self._cache_signature:
            return (
                self._cache_headers,
                self._cache_rows,
                self._cache_column_values,
            )

        wb = load_workbook(
            io.BytesIO(self._decrypted_bytes()), read_only=True, data_only=False
        )
        try:
            ws = self._sheet(wb)
            header_row = next(
                ws.iter_rows(min_row=1, max_row=1, values_only=True), ()
            )
            headers = [
                str(value).strip() if value is not None else ""
                for value in header_row
            ]
            if not headers:
                headers = list(DEFAULT_COLUMNS)

            name_idx = headers.index(REQUIRED_NAME) if REQUIRED_NAME in headers else None
            date_idx = headers.index(REQUIRED_DATE) if REQUIRED_DATE in headers else None
            collected = {header: set() for header in headers}
            records = []

            for row_number, values in enumerate(
                ws.iter_rows(min_row=2, values_only=True), start=2
            ):
                if not values:
                    continue

                row = {
                    header: values[index] if index < len(values) else None
                    for index, header in enumerate(headers)
                }
                non_rate = {
                    key: value
                    for key, value in row.items()
                    if key not in RATE_COLUMNS
                }
                if not any(
                    value is not None and str(value).strip()
                    for value in non_rate.values()
                ):
                    continue

                for index, header in enumerate(headers):
                    if header in RATE_COLUMNS or index >= len(values):
                        continue
                    value = values[index]
                    if value is not None:
                        text = str(value).strip()
                        if text:
                            collected[header].add(text)

                row["_row"] = row_number
                if name_idx is not None:
                    row[REQUIRED_NAME] = str(
                        row.get(REQUIRED_NAME, "") or ""
                    ).strip()
                if date_idx is not None:
                    row[REQUIRED_DATE] = (
                        normalize_date(row.get(REQUIRED_DATE)) or ""
                    )

                if "达成率" in headers and "合格率" in headers:
                    need_rate = (
                        row.get("达成率") is None
                        or str(row.get("达成率")).strip() in ("", "—")
                        or str(row.get("达成率")).strip().startswith("=")
                    )
                    need_quality = (
                        row.get("合格率") is None
                        or str(row.get("合格率")).strip() in ("", "—")
                        or str(row.get("合格率")).strip().startswith("=")
                    )
                    if need_rate or need_quality:
                        rate, quality = calc_rates(
                            row.get("当日计划产量"),
                            row.get("实际产量"),
                            row.get("合格品数量"),
                            row.get("不良品数量"),
                        )
                        if need_rate:
                            row["达成率"] = rate
                        if need_quality:
                            row["合格率"] = quality
                records.append(row)
        finally:
            wb.close()

        column_values = {
            header: tuple(sorted(values, key=lambda text: text.lower()))
            for header, values in collected.items()
        }
        self._cache_signature = signature
        self._cache_headers = tuple(headers)
        self._cache_rows = records
        self._cache_column_values = column_values
        return self._cache_headers, self._cache_rows, self._cache_column_values

    def _save_workbook(self, wb):
        buf = io.BytesIO()
        wb.calculation.fullCalcOnLoad = True
        wb.save(buf)
        payload = encrypt_bytes(self.secret, buf.getvalue())
        try:
            tmp = self.path + ".tmp"
            with open(tmp, "wb") as f:
                f.write(payload)
            os.replace(tmp, self.path)
            self._invalidate_cache()
        except PermissionError:
            raise AppError(
                "保存失败：数据文件正被其他程序占用，请稍后重试。"
            )

    def check_key(self):
        if os.path.exists(self.path):
            self._snapshot()
        else:
            self.ensure_file()

    def _sheet(self, wb):
        if "员工数据" in wb.sheetnames:
            return wb["员工数据"]
        return wb[wb.sheetnames[0]]

    def ensure_file(self):
        if not os.path.exists(self.path):
            wb = Workbook()
            ws = wb.active
            ws.title = "员工数据"
            ws.append(DEFAULT_COLUMNS)
            style_header_row(ws, len(DEFAULT_COLUMNS))
            self._fill_rate_formulas(ws, 2, RATE_FILL_ROWS)
            info = wb.create_sheet("使用说明")
            tips = [
                ["本表用于保存员工的每日数据（已加密保存）"],
                ["1. 数据表已加密，无法直接用 Excel 打开或编辑。"],
                ["2. 请通过程序的“添加”页面录入数据；达成率、合格率自动计算。"],
                ["3. 查询结果可通过“导出到Excel”正常查看。"],
                ["4. 每行代表一名员工某一天的一条记录；日期格式 2026-08-24。"],
                ["5. 可以自行增删其他列，程序会自动识别。"],
            ]
            for row in tips:
                info.append(row)
            info.column_dimensions["A"].width = 70
            self._save_workbook(wb)

    def _fill_rate_formulas(self, ws, start_row, end_row):
        """为整列预设达成率/合格率公式，方便直接在 Excel 中填写数量自动计算。"""
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        for r in range(start_row, end_row + 1):
            rate_f, quality_f = rate_formulas(r, headers)
            for i, header in enumerate(headers, 1):
                cell = ws.cell(row=r, column=i)
                if header == "达成率" and rate_f:
                    cell.value = rate_f
                elif header == "合格率" and quality_f:
                    cell.value = quality_f
                if Font is not None and header in RATE_COLUMNS:
                    cell.font = Font(color="D32F2F", bold=True, name="微软雅黑", size=11)

    def columns(self):
        headers, _, _ = self._snapshot()
        return list(headers)

    def names(self):
        result, seen = [], set()
        for row in self.rows():
            name = str(row.get(REQUIRED_NAME, "") or "").strip()
            if name and name not in seen:
                seen.add(name)
                result.append(name)
        return sorted(result, key=lambda s: s.lower())

    def teams(self):
        result, seen = [], set()
        for row in self.rows():
            team = str(row.get("班组名称", "") or "").strip()
            if team and team not in seen:
                seen.add(team)
                result.append(team)
        return sorted(result, key=lambda s: s.lower())

    def all_column_values(self):
        """返回每列出现过的所有不重复值，用于添加页的下拉列表。"""
        _, _, collected = self._snapshot()
        return {header: list(values) for header, values in collected.items()}

    def rows(
        self,
        name=None,
        team=None,
        emp_id=None,
        machine=None,
        material=None,
        customer=None,
        line=None,
        mold=None,
        newbie=None,
        support=None,
        start=None,
        end=None,
    ):
        _, records, _ = self._snapshot()
        data = []
        for cached_row in records:
            row = dict(cached_row)
            row_name = str(row.get(REQUIRED_NAME, "") or "").strip()
            row_date = row.get(REQUIRED_DATE, "") or ""
            row_team = str(row.get("班组名称", "") or "").strip()
            row_emp = str(row.get("工号", "") or "").strip()
            row_machine = str(row.get("所属机种", "") or "").strip()
            row_material = str(row.get("物料代码", "") or "").strip()
            row_customer = str(row.get("产品客户", "") or "").strip()
            row_line = str(row.get("所属产线/工序", "") or "").strip()
            row_mold = str(row.get("模具号", "") or "").strip()
            row_newbie = str(row.get("是否新人", "") or "").strip()
            row_support = str(row.get("是否支援人员", "") or "").strip()
            if name and row_name != str(name).strip():
                continue
            if team and row_team != str(team).strip():
                continue
            if emp_id and row_emp != str(emp_id).strip():
                continue
            if machine and row_machine != str(machine).strip():
                continue
            if material and row_material != str(material).strip():
                continue
            if customer and row_customer != str(customer).strip():
                continue
            if line and row_line != str(line).strip():
                continue
            if mold and row_mold != str(mold).strip():
                continue
            if newbie and row_newbie != "是":
                continue
            if support and row_support != "是":
                continue
            if start and (not row_date or row_date < start):
                continue
            if end and (not row_date or row_date > end):
                continue
            data.append(row)
        return data

    def add_row(self, record):
        wb = self._workbook()
        ws = self._sheet(wb)
        headers = self._headers_from_sheet(ws)
        name_idx = headers.index(REQUIRED_NAME) if REQUIRED_NAME in headers else 0
        target_row = None
        for r in range(2, ws.max_row + 2):
            cell = ws.cell(row=r, column=name_idx + 1)
            if cell.value is None or str(cell.value).strip() == "":
                target_row = r
                break
        if target_row is None:
            target_row = ws.max_row + 1
        for i, header in enumerate(headers, 1):
            ws.cell(row=target_row, column=i).value = record.get(header, "")
        rate_f, quality_f = rate_formulas(target_row, headers)
        if Font is not None:
            for i, header in enumerate(headers, 1):
                cell = ws.cell(row=target_row, column=i)
                if header in RATE_COLUMNS:
                    cell.font = Font(
                        color="D32F2F", bold=True, name="微软雅黑", size=11
                    )
                if header == "达成率" and rate_f:
                    cell.value = rate_f
                elif header == "合格率" and quality_f:
                    cell.value = quality_f
        self._save_workbook(wb)

    def add_rows_batch(self, records):
        """一次性写入多条记录，只解密和保存一次。"""
        if not records:
            return 0
        wb = self._workbook()
        ws = self._sheet(wb)
        headers = self._headers_from_sheet(ws)
        name_idx = headers.index(REQUIRED_NAME) if REQUIRED_NAME in headers else 0
        target_row = 2
        while (
            target_row <= ws.max_row
            and ws.cell(row=target_row, column=name_idx + 1).value is not None
            and str(
                ws.cell(row=target_row, column=name_idx + 1).value
            ).strip()
        ):
            target_row += 1
        for record in records:
            for column, header in enumerate(headers, 1):
                ws.cell(row=target_row, column=column).value = record.get(
                    header, ""
                )
            rate_f, quality_f = rate_formulas(target_row, headers)
            if Font is not None:
                for column, header in enumerate(headers, 1):
                    cell = ws.cell(row=target_row, column=column)
                    if header in RATE_COLUMNS:
                        cell.font = Font(
                            color="D32F2F",
                            bold=True,
                            name="微软雅黑",
                            size=11,
                        )
                    if header == "达成率" and rate_f:
                        cell.value = rate_f
                    elif header == "合格率" and quality_f:
                        cell.value = quality_f
            target_row += 1
        self._fill_rate_formulas(ws, 2, max(RATE_FILL_ROWS, target_row))
        self._save_workbook(wb)
        return len(records)

    def delete_rows(
        self, name=None, team=None, emp_id=None, start=None, end=None, preview=False
    ):
        """删除符合条件的记录；preview=True 时只返回条数，不实际删除。"""
        wb = self._workbook()
        ws = self._sheet(wb)
        headers = self._headers_from_sheet(ws)
        name_idx = headers.index(REQUIRED_NAME) if REQUIRED_NAME in headers else None
        date_idx = headers.index(REQUIRED_DATE) if REQUIRED_DATE in headers else None
        team_idx = headers.index("班组名称") if "班组名称" in headers else None
        emp_idx = headers.index("工号") if "工号" in headers else None
        targets = []
        for r in range(2, ws.max_row + 1):
            values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
            non_rate = [
                v for i, v in enumerate(values) if headers[i] not in RATE_COLUMNS
            ]
            if not any(v is not None and str(v).strip() for v in non_rate):
                continue
            row_name = (
                str(values[name_idx]).strip()
                if name_idx is not None and values[name_idx] is not None
                else ""
            )
            row_date = normalize_date(values[date_idx]) if date_idx is not None else None
            row_team = (
                str(values[team_idx]).strip()
                if team_idx is not None and values[team_idx] is not None
                else ""
            )
            row_emp = (
                str(values[emp_idx]).strip()
                if emp_idx is not None and values[emp_idx] is not None
                else ""
            )
            if name and row_name != str(name).strip():
                continue
            if team and row_team != str(team).strip():
                continue
            if emp_id and row_emp != str(emp_id).strip():
                continue
            if start and (not row_date or row_date < start):
                continue
            if end and (not row_date or row_date > end):
                continue
            targets.append(r)
        if preview:
            return len(targets)
        for r in reversed(targets):
            ws.delete_rows(r, 1)
        self._fill_rate_formulas(ws, 2, RATE_FILL_ROWS)
        self._save_workbook(wb)
        return len(targets)

    def delete_rows_by_indexes(self, indexes):
        """按 Excel 行号删除指定记录。"""
        wb = self._workbook()
        ws = self._sheet(wb)
        for idx in sorted(set(indexes), reverse=True):
            ws.delete_rows(idx, 1)
        self._fill_rate_formulas(ws, 2, RATE_FILL_ROWS)
        self._save_workbook(wb)
        return len(indexes)


class OptionStore:
    """在加密文件中维护查询和添加表单共用的下拉选项。"""

    def __init__(self, path, seed_values=None, legacy_dir=None):
        self.path = path
        self.revision = 0
        self._data = self._empty_data()
        if os.path.exists(path):
            self._load()
        else:
            self._seed(seed_values or {}, legacy_dir)
            self._save()

    @staticmethod
    def _empty_data():
        return {
            "version": 1,
            "initialized": True,
            "options": {field: [] for field in OPTION_FIELDS},
        }

    @staticmethod
    def _timestamp():
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _load(self):
        try:
            raw = open(self.path, "rb").read()
            payload = decrypt_bytes(PROGRAM_KEY, raw)
            data = json.loads(payload.decode("utf-8"))
            self._data = self._normalize(data)
        except Exception as exc:
            backup = self.path + ".bak"
            if not os.path.exists(backup):
                raise AppError(f"选项内容文件无法读取：{exc}")
            raw = open(backup, "rb").read()
            payload = decrypt_bytes(PROGRAM_KEY, raw)
            self._data = self._normalize(json.loads(payload.decode("utf-8")))
            self._save()

    def _normalize(self, data):
        result = self._empty_data()
        source = data.get("options", {}) if isinstance(data, dict) else {}
        for field in OPTION_FIELDS:
            entries = []
            seen = set()
            for item in source.get(field, []):
                if isinstance(item, str):
                    item = {"value": item, "active": True}
                value = str(item.get("value", "")).strip()
                key = value.casefold()
                if not value or key in seen:
                    continue
                seen.add(key)
                entries.append(
                    {
                        "value": value,
                        "active": bool(item.get("active", True)),
                        "created_at": str(
                            item.get("created_at") or self._timestamp()
                        ),
                    }
                )
            result["options"][field] = sorted(
                entries, key=lambda entry: entry["value"].casefold()
            )
        return result

    def _seed(self, seed_values, legacy_dir):
        seeded = {field: set() for field in OPTION_FIELDS}
        if legacy_dir:
            for field in OPTION_FIELDS:
                filename = OPTION_FILES.get(field)
                if not filename:
                    continue
                path = os.path.join(legacy_dir, filename)
                if not os.path.exists(path):
                    continue
                try:
                    with open(path, "r", encoding="utf-8-sig") as source:
                        for line in source:
                            value = line.strip()
                            if value:
                                seeded[field].add(value)
                except OSError:
                    pass
        for field in OPTION_FIELDS:
            for value in seed_values.get(field, []):
                text = str(value).strip()
                if text:
                    seeded[field].add(text)
        for field, values in seeded.items():
            self._data["options"][field] = [
                {"value": value, "active": True, "created_at": self._timestamp()}
                for value in sorted(values, key=str.casefold)
            ]

    def _save(self):
        payload = json.dumps(
            self._data, ensure_ascii=False, indent=2
        ).encode("utf-8")
        encrypted = encrypt_bytes(PROGRAM_KEY, payload)
        temporary = self.path + ".tmp"
        with open(temporary, "wb") as output:
            output.write(encrypted)
        if os.path.exists(self.path):
            import shutil

            shutil.copy2(self.path, self.path + ".bak")
        os.replace(temporary, self.path)
        self.revision += 1

    def values(self, field):
        return [
            entry["value"]
            for entry in self._data["options"].get(field, [])
            if entry["active"]
        ]

    def entries(self, field, include_inactive=False):
        return [
            dict(entry)
            for entry in self._data["options"].get(field, [])
            if include_inactive or entry["active"]
        ]

    def count(self, field):
        return len(self.values(field))

    def add(self, field, value):
        value = str(value).strip()
        if not value:
            raise AppError("选项内容不能为空。")
        entries = self._data["options"].setdefault(field, [])
        for entry in entries:
            if entry["value"].casefold() == value.casefold():
                if entry["active"]:
                    raise AppError(f"“{value}”已经存在。")
                entry["active"] = True
                entry["created_at"] = self._timestamp()
                self._save()
                return
        entries.append(
            {
                "value": value,
                "active": True,
                "created_at": self._timestamp(),
            }
        )
        entries.sort(key=lambda entry: entry["value"].casefold())
        self._save()

    def add_many(self, values_by_field):
        changed = False
        for field, values in (values_by_field or {}).items():
            if field not in OPTION_FIELDS:
                continue
            entries = self._data["options"].setdefault(field, [])
            by_key = {
                entry["value"].casefold(): entry for entry in entries
            }
            for raw_value in values:
                value = str(raw_value).strip()
                if not value:
                    continue
                key = value.casefold()
                entry = by_key.get(key)
                if entry is None:
                    entry = {
                        "value": value,
                        "active": True,
                        "created_at": self._timestamp(),
                    }
                    entries.append(entry)
                    by_key[key] = entry
                    changed = True
                elif not entry["active"]:
                    entry["active"] = True
                    entry["created_at"] = self._timestamp()
                    changed = True
            entries.sort(key=lambda entry: entry["value"].casefold())
        if changed:
            self._save()
        return changed

    def set_active(self, field, value, active):
        for entry in self._data["options"].get(field, []):
            if entry["value"] == value:
                if entry["active"] == bool(active):
                    return False
                entry["active"] = bool(active)
                self._save()
                return True
        return False

    def delete_permanently(self, field, values):
        targets = {str(value).strip() for value in values}
        entries = self._data["options"].get(field, [])
        remaining = [
            entry for entry in entries if entry["value"] not in targets
        ]
        removed = len(entries) - len(remaining)
        if removed:
            self._data["options"][field] = remaining
            self._save()
        return removed


class PillTabBar(tk.Canvas):
    """椭圆胶囊式选项卡按钮栏。"""

    def __init__(self, master, labels, command, background="#eef8f4"):
        super().__init__(
            master,
            height=54,
            background=background,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.labels = list(labels)
        self.command = command
        self.selected_index = 0
        self._boxes = []
        self.bind("<Configure>", lambda event: self._draw())
        self.bind("<Button-1>", self._on_click)
        self._draw()

    def select(self, index, notify=False):
        index = max(0, min(int(index), len(self.labels) - 1))
        changed = index != self.selected_index
        self.selected_index = index
        self._draw()
        if notify and changed:
            self.command(index)

    def _draw(self):
        self.delete("all")
        self._boxes = []
        x = 6
        height = 54
        y1, y2 = 8, height - 8
        radius = (y2 - y1) / 2
        for index, label in enumerate(self.labels):
            width = max(104, 26 + len(label) * 19)
            x2 = x + width
            selected = index == self.selected_index
            fill = "#dceff6" if selected else "#e8f6ee"
            outline = "#3f879d" if selected else "#c5dde1"
            points = [
                x + radius,
                y1,
                x2 - radius,
                y1,
                x2,
                y1,
                x2,
                y1 + radius,
                x2,
                y2 - radius,
                x2,
                y2,
                x2 - radius,
                y2,
                x + radius,
                y2,
                x,
                y2,
                x,
                y2 - radius,
                x,
                y1 + radius,
                x,
                y1,
            ]
            self.create_polygon(
                points,
                smooth=True,
                splinesteps=24,
                fill=fill,
                outline=outline,
                width=1,
            )
            self.create_text(
                (x + x2) / 2,
                (y1 + y2) / 2,
                text=label,
                fill="#27434b",
                font=(FONT_FAMILY, 11, "bold"),
            )
            self._boxes.append((x, x2, y1, y2))
            x = x2 + 8

    def _on_click(self, event):
        for index, (x1, x2, y1, y2) in enumerate(self._boxes):
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                self.select(index, notify=True)
                return


class BasePage(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        raise NotImplementedError

    def on_show(self):
        pass


class QueryPage(BasePage):
    def _build(self):
        self._options_revision = None
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew", padx=40, pady=(22, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="员工数据查询", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="按条件组合查询每日记录，日期可不填",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Button(header, text="刷新数据", command=self.refresh_data).grid(
            row=0, column=1, rowspan=2, sticky="e", padx=(24, 0)
        )

        card_frame = tk.Frame(
            self, bg="#e8f5fb", highlightbackground="#c8e1e2", highlightthickness=1
        )
        card_frame.grid(row=1, column=0, sticky="ew", padx=40, pady=10)
        card = ttk.Frame(card_frame, style="BlueSection.TFrame", padding=(20, 16))
        card.pack(fill="both", expand=True)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(4, weight=1)

        ttk.Label(
            card, text="基本查询", style="BlueTitle.TLabel"
        ).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 4))

        ttk.Label(card, text="员工姓名：", style="Field.TLabel").grid(
            row=1, column=0, sticky="e", padx=(0, 12), pady=8
        )
        self.name_var = tk.StringVar()
        self.name_combo = FilterCombobox(
            card,
            textvariable=self.name_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.name_combo.grid(row=1, column=1, sticky="ew", pady=8)

        ttk.Label(card, text="工号：", style="Field.TLabel").grid(
            row=1, column=3, sticky="e", padx=(32, 12), pady=8
        )
        self.emp_id_var = tk.StringVar()
        self.emp_id_combo = FilterCombobox(
            card,
            textvariable=self.emp_id_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.emp_id_combo.grid(row=1, column=4, sticky="ew", pady=8)

        ttk.Label(card, text="班组名称：", style="Field.TLabel").grid(
            row=2, column=0, sticky="e", padx=(0, 12), pady=8
        )
        self.team_var = tk.StringVar()
        self.team_combo = FilterCombobox(
            card,
            textvariable=self.team_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.team_combo.grid(row=2, column=1, sticky="ew", pady=8)

        ttk.Label(card, text="所属机种：", style="Field.TLabel").grid(
            row=2, column=3, sticky="e", padx=(32, 12), pady=8
        )
        self.machine_var = tk.StringVar()
        self.machine_combo = FilterCombobox(
            card,
            textvariable=self.machine_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.machine_combo.grid(row=2, column=4, sticky="ew", pady=8)

        ttk.Label(card, text="物料代码：", style="Field.TLabel").grid(
            row=3, column=0, sticky="e", padx=(0, 12), pady=8
        )
        self.material_var = tk.StringVar()
        self.material_combo = FilterCombobox(
            card,
            textvariable=self.material_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.material_combo.grid(row=3, column=1, sticky="ew", pady=8)

        ttk.Label(card, text="产品客户：", style="Field.TLabel").grid(
            row=3, column=3, sticky="e", padx=(32, 12), pady=8
        )
        self.customer_var = tk.StringVar()
        self.customer_combo = FilterCombobox(
            card,
            textvariable=self.customer_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.customer_combo.grid(row=3, column=4, sticky="ew", pady=8)

        ttk.Label(card, text="所属产线/工序：", style="Field.TLabel").grid(
            row=4, column=0, sticky="e", padx=(0, 12), pady=8
        )
        self.line_var = tk.StringVar()
        self.line_combo = FilterCombobox(
            card,
            textvariable=self.line_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.line_combo.grid(row=4, column=1, sticky="ew", pady=8)

        ttk.Label(card, text="模具号：", style="Field.TLabel").grid(
            row=4, column=3, sticky="e", padx=(32, 12), pady=8
        )
        self.mold_var = tk.StringVar()
        self.mold_combo = FilterCombobox(
            card,
            textvariable=self.mold_var,
            font=(FONT_FAMILY, 12),
            width=22,
            state="readonly",
        )
        self.mold_combo.grid(row=4, column=4, sticky="ew", pady=8)

        self.filter_enabled = tk.BooleanVar(value=True)
        date_sections = ttk.Frame(card, style="BlueSection.TFrame")
        date_sections.grid(
            row=5, column=0, columnspan=5, sticky="ew", pady=(10, 5)
        )
        date_sections.columnconfigure(0, weight=1, uniform="date_sections")
        date_sections.columnconfigure(1, weight=1, uniform="date_sections")
        filter_section = ttk.Frame(
            date_sections, style="GreenSection.TFrame", padding=(14, 10)
        )
        filter_section.grid(
            row=0, column=0, sticky="nsew", padx=(0, 5)
        )
        filter_section.columnconfigure(1, weight=1)
        ttk.Label(
            filter_section,
            text="日期筛选",
            style="GreenTitle.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        filter_box = ttk.Frame(filter_section, style="GreenSection.TFrame")
        filter_box.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 2))
        ToggleSwitch(
            filter_box,
            self.filter_enabled,
            command=self._toggle_filter,
            bg="#edf8f1",
        ).pack(side="left")
        ttk.Label(
            filter_box, text="启用日期筛选", style="GreenField.TLabel"
        ).pack(side="left", padx=(6, 0))
        ttk.Label(
            filter_box,
            text="点击日历按钮选择",
            style="GreenSub.TLabel",
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            filter_section, text="开始日期：", style="GreenField.TLabel"
        ).grid(
            row=2, column=0, sticky="e", padx=(0, 10), pady=4
        )
        self.start_entry = ClearableDateEntry(
            filter_section,
            date_pattern="yyyy-mm-dd",
            locale="zh_CN",
            state="disabled",
            width=14,
        )
        self.start_entry.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(
            filter_section, text="结束日期：", style="GreenField.TLabel"
        ).grid(
            row=3, column=0, sticky="e", padx=(0, 10), pady=4
        )
        self.end_entry = ClearableDateEntry(
            filter_section,
            date_pattern="yyyy-mm-dd",
            locale="zh_CN",
            state="disabled",
            width=14,
        )
        self.end_entry.grid(row=3, column=1, sticky="w", pady=4)

        date_query_section = ttk.Frame(
            date_sections, style="BlueQuerySection.TFrame", padding=(14, 10)
        )
        date_query_section.grid(
            row=0, column=1, sticky="nsew", padx=(5, 0)
        )
        date_query_section.columnconfigure(1, weight=1)
        ttk.Label(
            date_query_section,
            text="日期查询",
            style="BlueQueryTitle.TLabel",
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            date_query_section,
            text="指定日期：",
            style="BlueQueryField.TLabel",
        ).grid(row=1, column=0, sticky="e", padx=(0, 10), pady=(6, 2))
        self.single_date_entry = ClearableDateEntry(
            date_query_section,
            date_pattern="yyyy-mm-dd",
            locale="zh_CN",
            width=14,
        )
        self.single_date_entry.grid(row=1, column=1, sticky="w", pady=(6, 2))
        self.single_date_entry.set_date(datetime.date.today())
        ttk.Button(
            date_query_section,
            text="日期查询",
            style="Primary.TButton",
            command=self.query_date,
        ).grid(row=1, column=2, sticky="w", padx=(12, 0), pady=(6, 2))

        buttons = ttk.Frame(card, style="BlueSection.TFrame")
        buttons.grid(
            row=6, column=0, columnspan=5, sticky="ew", pady=(10, 0)
        )
        button_specs = (
            ("查询", self.do_query, "Primary.TButton"),
            ("新人查询", self.query_newbie, "TButton"),
            ("支援查询", self.query_support, "TButton"),
            ("清空条件", self.clear_filters, "TButton"),
        )
        for column, (text, command, button_style) in enumerate(button_specs):
            buttons.columnconfigure(column, weight=1, uniform="query_actions")
            ttk.Button(
                buttons,
                text=text,
                style=button_style,
                command=command,
            ).grid(
                row=0,
                column=column,
                sticky="ew",
                padx=(0 if column == 0 else 4, 0 if column == 3 else 4),
            )

        ttk.Label(
            self,
            text="提示：数据表已加密保存，无法直接用 Excel 打开；查询结果可通过“导出到Excel”查看。",
            style="Sub.TLabel",
        ).grid(row=2, column=0, sticky="sw", padx=38, pady=(2, 0))
        self.refresh_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.refresh_var, style="Sub.TLabel").grid(
            row=2, column=0, sticky="se", padx=38, pady=(2, 0)
        )

    def on_show(self):
        if self._options_revision != self.app.options.revision:
            self._refresh_filters()
        self._toggle_filter()
        self.name_combo.focus_set()

    def _refresh_filters(self):
        options = self.app.options
        set_combo_values(self.name_combo, options.values("员工姓名"))
        set_combo_values(self.team_combo, options.values("班组名称"))
        set_combo_values(self.emp_id_combo, options.values("工号"))
        set_combo_values(self.machine_combo, options.values("所属机种"))
        set_combo_values(self.material_combo, options.values("物料代码"))
        set_combo_values(self.customer_combo, options.values("产品客户"))
        set_combo_values(self.line_combo, options.values("所属产线/工序"))
        set_combo_values(self.mold_combo, options.values("模具号"))
        self._options_revision = options.revision

    def refresh_data(self):
        try:
            self._refresh_filters()
        except Exception:
            messagebox.showerror(
                "刷新失败",
                "无法读取“员工数据.xlsx”，请确认文件未被占用后重试。",
            )
            return
        self.refresh_var.set("已刷新 " + datetime.datetime.now().strftime("%H:%M:%S"))

    def _toggle_filter(self):
        state = "normal" if self.filter_enabled.get() else "disabled"
        self.start_entry.config(state=state)
        self.end_entry.config(state=state)

    def clear_filters(self):
        self.name_var.set("")
        self.team_var.set("")
        self.emp_id_var.set("")
        self.machine_var.set("")
        self.material_var.set("")
        self.customer_var.set("")
        self.line_var.set("")
        self.mold_var.set("")
        self.start_entry.delete(0, "end")
        self.end_entry.delete(0, "end")
        self.single_date_entry.delete(0, "end")

    def do_query(self):
        data = self.app.data
        headers = data.columns()
        if REQUIRED_NAME not in headers:
            messagebox.showerror("提示", "数据表中找不到“员工姓名”列，请保留该列后重试。")
            return
        name = self.name_var.get().strip()
        team = self.team_var.get().strip()
        emp_id = self.emp_id_var.get().strip()
        machine = self.machine_var.get().strip()
        material = self.material_var.get().strip()
        customer = self.customer_var.get().strip()
        line = self.line_var.get().strip()
        mold = self.mold_var.get().strip()
        if not any((name, team, emp_id, machine, material, customer, line, mold)):
            messagebox.showwarning("提示", "请至少填写一项查询条件。")
            return
        start = end = None
        if self.filter_enabled.get():
            start_date = get_date_or_none(self.start_entry)
            end_date = get_date_or_none(self.end_entry)
            start = start_date.isoformat() if start_date else None
            end = end_date.isoformat() if end_date else None
            if start and end and start > end:
                messagebox.showwarning("提示", "开始日期不能晚于结束日期。")
                return
        rows = data.rows(
            name=name or None,
            team=team or None,
            emp_id=emp_id or None,
            machine=machine or None,
            material=material or None,
            customer=customer or None,
            line=line or None,
            mold=mold or None,
            start=start,
            end=end,
        )
        if name:
            criteria = []
            if team:
                criteria.append(f"班组：{team}")
            if emp_id:
                criteria.append(f"工号：{emp_id}")
            if machine:
                criteria.append(f"所属机种：{machine}")
            if material:
                criteria.append(f"物料代码：{material}")
            if customer:
                criteria.append(f"产品客户：{customer}")
            if line:
                criteria.append(f"所属产线/工序：{line}")
            if mold:
                criteria.append(f"模具号：{mold}")
            if start or end:
                criteria.append(f"日期：{start or '不限'} ~ {end or '不限'}")
            self.app.pages["result"].load(
                name,
                rows,
                start,
                end,
                criteria_text=("筛选条件：" + "；".join(criteria)) if criteria else None,
            )
            self.app.show("result")
            return
        if team:
            if "班组名称" not in headers:
                messagebox.showerror("提示", "数据表中找不到“班组名称”列，请保留该列后重试。")
                return
            if not rows:
                messagebox.showinfo("提示", "该班组暂无符合条件的记录。")
                return
            self.app.pages["team"].load(team, rows, start, end)
            self.app.show("team")
            return
        criteria = []
        if emp_id:
            criteria.append(f"工号：{emp_id}")
        if machine:
            criteria.append(f"所属机种：{machine}")
        if material:
            criteria.append(f"物料代码：{material}")
        if customer:
            criteria.append(f"产品客户：{customer}")
        if line:
            criteria.append(f"所属产线/工序：{line}")
        if mold:
            criteria.append(f"模具号：{mold}")
        if start or end:
            criteria.append(f"日期：{start or '不限'} ~ {end or '不限'}")
        self.app.pages["result"].load(
            "", rows, start, end, criteria_text="筛选条件：" + "；".join(criteria)
        )
        self.app.show("result")

    def query_newbie(self):
        data = self.app.data
        headers = data.columns()
        if "是否新人" not in headers:
            messagebox.showerror("提示", "数据表中找不到“是否新人”列，请保留该列后重试。")
            return
        start = end = None
        if self.filter_enabled.get():
            start_date = get_date_or_none(self.start_entry)
            end_date = get_date_or_none(self.end_entry)
            start = start_date.isoformat() if start_date else None
            end = end_date.isoformat() if end_date else None
            if start and end and start > end:
                messagebox.showwarning("提示", "开始日期不能晚于结束日期。")
                return
        rows = data.rows(newbie="是", start=start, end=end)
        if not rows:
            messagebox.showinfo("提示", "暂无符合条件的“新人”记录。")
            return
        criteria = ["新人（是否新人=是）"]
        if start or end:
            criteria.append(f"日期：{start or '不限'} ~ {end or '不限'}")
        self.app.pages["result"].load(
            "", rows, start, end, criteria_text="筛选条件：" + "；".join(criteria)
        )
        self.app.show("result")

    def query_support(self):
        data = self.app.data
        headers = data.columns()
        if "是否支援人员" not in headers:
            messagebox.showerror(
                "提示", "数据表中找不到“是否支援人员”列，请保留该列后重试。"
            )
            return
        start = end = None
        if self.filter_enabled.get():
            start_date = get_date_or_none(self.start_entry)
            end_date = get_date_or_none(self.end_entry)
            start = start_date.isoformat() if start_date else None
            end = end_date.isoformat() if end_date else None
            if start and end and start > end:
                messagebox.showwarning("提示", "开始日期不能晚于结束日期。")
                return
        rows = data.rows(support="是", start=start, end=end)
        if not rows:
            messagebox.showinfo("提示", "暂无符合条件的“支援人员”记录。")
            return
        criteria = ["支援人员（是否支援人员=是）"]
        if start or end:
            criteria.append(f"日期：{start or '不限'} ~ {end or '不限'}")
        self.app.pages["result"].load(
            "", rows, start, end, criteria_text="筛选条件：" + "；".join(criteria)
        )
        self.app.show("result")

    def query_date(self):
        selected = get_date_or_none(self.single_date_entry)
        if selected is None:
            messagebox.showwarning("提示", "请选择要查询的日期。")
            return
        date_text = selected.isoformat()
        rows = self.app.data.rows(start=date_text, end=date_text)
        if not rows:
            messagebox.showinfo("提示", f"{date_text} 暂无员工数据记录。")
            return
        self.app.pages["result"].load(
            "",
            rows,
            date_text,
            date_text,
            criteria_text=f"指定日期：{date_text}",
        )
        self.app.show("result")

def export_chart_analysis(rows, columns, label):
    """导出热力图、重点趋势和原始数据到 Excel。"""
    if not rows:
        messagebox.showinfo("提示", "当前没有可导出的记录。")
        return
    safe = re.sub(r'[\\/:*?"<>|]', "_", str(label or "查询结果"))
    out_dir = os.path.join(base_dir(), "查询结果")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(
        out_dir,
        "图表分析_{0}_{1:%Y%m%d_%H%M%S}.xlsx".format(
            safe, datetime.datetime.now()
        ),
    )

    rows_sorted = sorted(
        rows,
        key=lambda row: (
            normalize_date(row.get(REQUIRED_DATE)) or "",
            str(row.get(REQUIRED_NAME, "") or ""),
        ),
    )
    wb = Workbook()
    raw = wb.active
    raw.title = "原始数据"
    raw.append(columns)
    style_header_row(raw, len(columns))
    for row in rows_sorted:
        raw.append([row.get(column, "") for column in columns])
    total_values, _ = summarize_rows(rows, columns)
    raw.append(total_values)
    total_fill = PatternFill("solid", fgColor="DCEBFF")
    for cell in raw[raw.max_row]:
        cell.fill = total_fill
        cell.font = Font(bold=True, color="27434B", name="微软雅黑", size=11)
    for index, column in enumerate(columns, 1):
        longest = max(
            [len(str(column))]
            + [len(str(row.get(column, ""))) for row in rows_sorted]
        )
        raw.column_dimensions[get_column_letter(index)].width = min(
            max(12, longest * 2 + 3), 60
        )
    rate_indexes = [
        index
        for index, column in enumerate(columns, 1)
        if column in RATE_COLUMNS
    ]
    for row_index in range(2, raw.max_row + 1):
        for column_index in rate_indexes:
            raw.cell(row_index, column_index).font = Font(
                color="D32F2F", bold=True, name="微软雅黑", size=11
            )
    raw.freeze_panes = "A2"
    raw.auto_filter.ref = (
        f"A1:{get_column_letter(len(columns))}{raw.max_row}"
    )

    grouped = {}
    dates = set()
    employees = set()
    for row in rows_sorted:
        name = str(row.get(REQUIRED_NAME, "") or "").strip()
        date = normalize_date(row.get(REQUIRED_DATE))
        if not name or not date:
            continue
        employees.add(name)
        dates.add(date)
        bucket = grouped.setdefault(
            (name, date),
            {"plan": 0.0, "actual": 0.0, "good": 0.0, "bad": 0.0},
        )
        bucket["plan"] += to_num(row.get("当日计划产量")) or 0
        bucket["actual"] += to_num(row.get("实际产量")) or 0
        bucket["good"] += to_num(row.get("合格品数量")) or 0
        bucket["bad"] += to_num(row.get("不良品数量")) or 0

    date_list = sorted(dates)
    employee_list = sorted(employees, key=str.casefold)
    heatmap = wb.create_sheet("热力图")
    header_fill = PatternFill("solid", fgColor="5A9CAF")
    header_font = Font(color="FFFFFF", bold=True, name="微软雅黑", size=11)

    def write_heatmap_block(start_row, title, metric):
        heatmap.cell(start_row, 1, f"{title}（%）")
        heatmap.cell(start_row, 1).font = Font(
            bold=True, color="2F6F82", name="微软雅黑", size=13
        )
        headers = ["员工姓名"] + date_list
        for column, value in enumerate(headers, 1):
            cell = heatmap.cell(start_row + 1, column, value)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        data_start = start_row + 2
        for row_offset, employee in enumerate(employee_list):
            heatmap.cell(data_start + row_offset, 1, employee)
            for column_offset, date in enumerate(date_list, 2):
                values = grouped.get((employee, date))
                ratio = None
                if values:
                    if metric == "达成率" and values["plan"] > 0:
                        ratio = values["actual"] / values["plan"]
                    elif metric == "合格率":
                        denominator = (
                            values["actual"]
                            if values["actual"] > 0
                            else values["good"] + values["bad"]
                        )
                        if denominator > 0:
                            ratio = values["good"] / denominator
                cell = heatmap.cell(data_start + row_offset, column_offset)
                if ratio is not None:
                    cell.value = ratio
                    cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="center")
        data_end = data_start + max(0, len(employee_list)) - 1
        if date_list and employee_list:
            first = f"B{data_start}"
            last = (
                f"{get_column_letter(len(date_list) + 1)}"
                f"{data_end}"
            )
            heatmap.conditional_formatting.add(
                f"{first}:{last}",
                ColorScaleRule(
                    start_type="num",
                    start_value=0.8,
                    start_color="F8696B",
                    mid_type="num",
                    mid_value=1.0,
                    mid_color="FFEB84",
                    end_type="num",
                    end_value=1.1,
                    end_color="63BE7B",
                ),
            )
        return data_end + 3

    next_row = write_heatmap_block(1, "达成率热力图", "达成率")
    write_heatmap_block(next_row, "合格率热力图", "合格率")
    heatmap.column_dimensions["A"].width = 18
    for index in range(2, len(date_list) + 2):
        heatmap.column_dimensions[get_column_letter(index)].width = 13
    heatmap.freeze_panes = "B3"

    by_date = {}
    for row in rows_sorted:
        date = normalize_date(row.get(REQUIRED_DATE))
        if not date:
            continue
        bucket = by_date.setdefault(
            date,
            {"plan": 0.0, "actual": 0.0, "good": 0.0, "bad": 0.0, "hours": 0.0},
        )
        bucket["plan"] += to_num(row.get("当日计划产量")) or 0
        bucket["actual"] += to_num(row.get("实际产量")) or 0
        bucket["good"] += to_num(row.get("合格品数量")) or 0
        bucket["bad"] += to_num(row.get("不良品数量")) or 0
        bucket["hours"] += to_num(row.get("工时")) or 0

    trend = wb.create_sheet("重点趋势")
    trend.append(["日期", "计划产量", "实际产量", "达成率", "合格率", "工时"])
    style_header_row(trend, 6)
    for date in sorted(by_date):
        values = by_date[date]
        achievement = (
            values["actual"] / values["plan"] if values["plan"] > 0 else None
        )
        quality = (
            values["good"] / values["actual"]
            if values["actual"] > 0
            else None
        )
        trend.append(
            [
                date,
                values["plan"],
                values["actual"],
                achievement,
                quality,
                values["hours"],
            ]
        )
        row_index = trend.max_row
        trend.cell(row_index, 4).number_format = "0.0%"
        trend.cell(row_index, 5).number_format = "0.0%"
    for column, width in (("A", 14), ("B", 14), ("C", 14), ("D", 12), ("E", 12), ("F", 10)):
        trend.column_dimensions[column].width = width
    if trend.max_row > 1:
        categories = Reference(
            trend, min_col=1, min_row=2, max_row=trend.max_row
        )
        quantity_chart = LineChart()
        quantity_chart.title = "计划与实际产量趋势"
        quantity_chart.y_axis.title = "产量"
        quantity_chart.x_axis.title = "日期"
        quantity_chart.height = 8
        quantity_chart.width = 17
        quantity_data = Reference(
            trend, min_col=2, max_col=3, min_row=1, max_row=trend.max_row
        )
        quantity_chart.add_data(quantity_data, titles_from_data=True)
        quantity_chart.set_categories(categories)
        trend.add_chart(quantity_chart, "H2")

        rate_chart = LineChart()
        rate_chart.title = "达成率与合格率趋势"
        rate_chart.y_axis.title = "百分比"
        rate_chart.x_axis.title = "日期"
        rate_chart.y_axis.numFmt = "0%"
        rate_chart.height = 8
        rate_chart.width = 17
        rate_data = Reference(
            trend, min_col=4, max_col=5, min_row=1, max_row=trend.max_row
        )
        rate_chart.add_data(rate_data, titles_from_data=True)
        rate_chart.set_categories(categories)
        trend.add_chart(rate_chart, "H22")

    wb.save(path)
    if messagebox.askyesno(
        "导出成功", f"图表分析已导出到：\n{path}\n\n是否现在打开？"
    ):
        try:
            os.startfile(path)
        except OSError:
            pass


class ResultPage(BasePage):
    def _build(self):
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(36, 10))
        self.title_var = tk.StringVar(value="查询结果")
        ttk.Label(header, textvariable=self.title_var, style="Title.TLabel").pack(anchor="w")
        self.subtitle_var = tk.StringVar(value="")
        ttk.Label(header, textvariable=self.subtitle_var, style="Sub.TLabel").pack(
            anchor="w", pady=(4, 0)
        )

        bar = ttk.Frame(self)
        bar.grid(row=1, column=0, sticky="ew", padx=36, pady=(0, 10))
        ttk.Button(bar, text="← 返回查询", command=lambda: self.app.show("query")).pack(side="left")
        ttk.Button(bar, text="导出到Excel", command=self.export).pack(side="left", padx=(12, 0))
        ttk.Button(
            bar, text="导出图表分析", command=self.export_analysis
        ).pack(side="left", padx=(12, 0))
        self.count_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count_var, style="Sub.TLabel").pack(side="right")

        box = ttk.Frame(self)
        box.grid(row=2, column=0, sticky="nsew", padx=36, pady=(0, 20))
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(box, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        summary = ttk.Frame(self, style="Card.TFrame", padding=(20, 14))
        summary.grid(row=3, column=0, sticky="ew", padx=36, pady=(0, 20))
        self.summary_var = tk.StringVar(value="")
        ttk.Label(
            summary,
            textvariable=self.summary_var,
            font=(FONT_FAMILY, 12, "bold"),
            foreground="#1f4e9c",
            justify="left",
        ).pack(anchor="w")

        self.rows_data = []
        self.columns = []
        self.current_name = ""

    def load(self, name, rows, start, end, criteria_text=None):
        self.current_name = name
        self.rows_data = rows
        self.columns = self.app.data.columns()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.columns
        for col in self.columns:
            if col in RATE_COLUMNS:
                try:
                    self.tree.heading(
                        col,
                        text=col,
                        foreground="#ffffff",
                        anchor="center",
                    )
                except tk.TclError:
                    self.tree.heading(col, text=col)
            else:
                self.tree.heading(col, text=col, anchor="center")
            width = RESULT_COLUMN_WIDTHS.get(col, 120)
            for r in rows:
                width = max(width, 60 + len(str(r.get(col, ""))) * 18)
            self.tree.column(
                col,
                width=min(width, 360),
                minwidth=RESULT_COLUMN_WIDTHS.get(col, 100),
                anchor=(
                    "w"
                    if col
                    in (REQUIRED_NAME, "班组名称", "产品客户", "所属机种", "物料代码", "所属产线/工序", "备注")
                    else "center"
                ),
                stretch=True,
            )
        rows_sorted = sorted(
            rows,
            key=lambda r: (
                str(r.get(REQUIRED_DATE, "") or ""),
                str(r.get(REQUIRED_NAME, "") or ""),
            ),
        )
        for idx, r in enumerate(rows_sorted):
            values = [r.get(c, "") for c in self.columns]
            self.tree.insert("", "end", values=values, tags=("even" if idx % 2 == 0 else "odd",))
        self.tree.tag_configure("odd", background="#f5f7fb")
        self.tree.tag_configure("even", background="#ffffff")

        total_values, totals = summarize_rows(rows, self.columns)
        self.tree.tag_configure("total", background="#dcebff")
        self.tree.insert("", "end", values=total_values, tags=("total",))

        date_text = ""
        if start or end:
            date_text = f"，日期范围：{start or '不限'} ~ {end or '不限'}"
        if name:
            self.title_var.set(f"员工“{name}”的记录")
        else:
            self.title_var.set("查询结果")
        if criteria_text:
            self.subtitle_var.set(f"{criteria_text}（数据来自“员工数据.xlsx”）")
        else:
            self.subtitle_var.set(f"以下内容来自数据表“员工数据.xlsx”{date_text}")
        self.count_var.set(f"共 {len(rows)} 条记录")
        self.summary_var.set(build_summary_text(totals))

    def export(self):
        if not self.rows_data:
            messagebox.showinfo("提示", "当前没有可导出的记录。")
            return
        safe = re.sub(r'[\\/:*?"<>|]', "_", self.current_name) or "结果"
        out_dir = os.path.join(base_dir(), "查询结果")
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(
            out_dir, "查询结果_{0}_{1:%Y%m%d}.xlsx".format(safe, datetime.date.today())
        )
        wb = Workbook()
        ws = wb.active
        ws.title = safe[:31]
        ws.append(self.columns)
        style_header_row(ws, len(self.columns))
        for r in sorted(self.rows_data, key=lambda x: str(x.get(REQUIRED_DATE, "") or "")):
            ws.append([r.get(c, "") for c in self.columns])
        for i, col in enumerate(self.columns, 1):
            max_len = len(str(col))
            for r in self.rows_data:
                max_len = max(max_len, len(str(r.get(col, ""))))
            ws.column_dimensions[get_column_letter(i)].width = min(max_len * 2 + 4, 60)
        if Font is not None:
            rate_indices = [i for i, col in enumerate(self.columns, 1) if col in RATE_COLUMNS]
            for r_idx in range(2, ws.max_row + 1):
                for c_idx in rate_indices:
                    ws.cell(row=r_idx, column=c_idx).font = Font(
                        color="D32F2F", bold=True, name="微软雅黑", size=11
                    )
        wb.save(path)
        if messagebox.askyesno("导出成功", f"已导出到：\n{path}\n\n是否现在打开？"):
            try:
                os.startfile(path)
            except OSError:
                pass

    def export_analysis(self):
        export_chart_analysis(
            self.rows_data,
            self.columns,
            self.current_name or "查询结果",
        )


class TeamResultPage(BasePage):
    """班组查询结果页：显示该班组所有成员的记录，并在底部汇总。"""

    def _build(self):
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(36, 10))
        self.title_var = tk.StringVar(value="班组查询结果")
        ttk.Label(header, textvariable=self.title_var, style="Title.TLabel").pack(anchor="w")
        self.subtitle_var = tk.StringVar(value="")
        ttk.Label(header, textvariable=self.subtitle_var, style="Sub.TLabel").pack(
            anchor="w", pady=(4, 0)
        )

        bar = ttk.Frame(self)
        bar.grid(row=1, column=0, sticky="ew", padx=36, pady=(0, 10))
        ttk.Button(bar, text="← 返回查询", command=lambda: self.app.show("query")).pack(side="left")
        ttk.Button(bar, text="导出到Excel", command=self.export).pack(side="left", padx=(12, 0))
        ttk.Button(
            bar, text="导出图表分析", command=self.export_analysis
        ).pack(side="left", padx=(12, 0))
        self.count_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count_var, style="Sub.TLabel").pack(side="right")

        box = ttk.Frame(self)
        box.grid(row=2, column=0, sticky="nsew", padx=36, pady=(0, 6))
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(box, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        summary = ttk.Frame(self, style="Card.TFrame", padding=(20, 14))
        summary.grid(row=3, column=0, sticky="ew", padx=36, pady=(0, 20))
        self.summary_var = tk.StringVar(value="")
        ttk.Label(
            summary,
            textvariable=self.summary_var,
            font=(FONT_FAMILY, 12, "bold"),
            foreground="#1f4e9c",
            justify="left",
        ).pack(anchor="w")

        self.rows_data = []
        self.columns = []
        self.current_team = ""

    def load(self, team, rows, start, end):
        self.current_team = team
        self.rows_data = rows
        self.columns = self.app.data.columns()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.columns
        for col in self.columns:
            if col in RATE_COLUMNS:
                try:
                    self.tree.heading(
                        col,
                        text=col,
                        foreground="#ffffff",
                        anchor="center",
                    )
                except tk.TclError:
                    self.tree.heading(col, text=col)
            else:
                self.tree.heading(col, text=col, anchor="center")
            width = RESULT_COLUMN_WIDTHS.get(col, 120)
            for r in rows:
                width = max(width, 60 + len(str(r.get(col, ""))) * 18)
            self.tree.column(
                col,
                width=min(width, 360),
                minwidth=RESULT_COLUMN_WIDTHS.get(col, 100),
                anchor=(
                    "w"
                    if col
                    in (REQUIRED_NAME, "班组名称", "产品客户", "所属机种", "物料代码", "所属产线/工序", "备注")
                    else "center"
                ),
                stretch=True,
            )

        rows_sorted = sorted(
            rows,
            key=lambda r: (
                str(r.get(REQUIRED_DATE, "") or ""),
                str(r.get(REQUIRED_NAME, "") or ""),
            ),
        )
        for idx, r in enumerate(rows_sorted):
            values = [r.get(c, "") for c in self.columns]
            self.tree.insert("", "end", values=values, tags=("even" if idx % 2 == 0 else "odd",))
        self.tree.tag_configure("odd", background="#f5f7fb")
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("total", background="#dcebff")

        total_values, totals = summarize_rows(rows, self.columns)
        self.tree.insert("", "end", values=total_values, tags=("total",))

        members = {
            str(r.get(REQUIRED_NAME, "") or "").strip()
            for r in rows
            if str(r.get(REQUIRED_NAME, "") or "").strip()
        }
        date_text = ""
        if start or end:
            date_text = f"，日期范围：{start or '不限'} ~ {end or '不限'}"
        self.title_var.set(f"班组“{team}”的记录")
        self.subtitle_var.set(f"以下为该班组所有成员在数据表中的记录{date_text}")
        self.count_var.set(f"共 {len(rows)} 条记录，{len(members)} 名成员")
        self.summary_var.set(build_summary_text(totals))

    def export(self):
        if not self.rows_data:
            messagebox.showinfo("提示", "当前没有可导出的记录。")
            return
        safe = re.sub(r'[\\/:*?"<>|]', "_", self.current_team) or "结果"
        out_dir = os.path.join(base_dir(), "查询结果")
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(
            out_dir, "班组结果_{0}_{1:%Y%m%d}.xlsx".format(safe, datetime.date.today())
        )
        wb = Workbook()
        ws = wb.active
        ws.title = safe[:31]
        ws.append(self.columns)
        style_header_row(ws, len(self.columns))
        for r in sorted(self.rows_data, key=lambda x: str(x.get(REQUIRED_DATE, "") or "")):
            ws.append([r.get(c, "") for c in self.columns])
        for i, col in enumerate(self.columns, 1):
            max_len = len(str(col))
            for r in self.rows_data:
                max_len = max(max_len, len(str(r.get(col, ""))))
            ws.column_dimensions[get_column_letter(i)].width = min(max_len * 2 + 4, 60)
        if Font is not None:
            rate_indices = [i for i, col in enumerate(self.columns, 1) if col in RATE_COLUMNS]
            for r_idx in range(2, ws.max_row + 1):
                for c_idx in rate_indices:
                    ws.cell(row=r_idx, column=c_idx).font = Font(
                        color="D32F2F", bold=True, name="微软雅黑", size=11
                    )
        wb.save(path)
        if messagebox.askyesno("导出成功", f"已导出到：\n{path}\n\n是否现在打开？"):
            try:
                os.startfile(path)
            except OSError:
                pass

    def export_analysis(self):
        export_chart_analysis(
            self.rows_data,
            self.columns,
            self.current_team or "班组结果",
        )


class AddPage(BasePage):
    def _build(self):
        self._page_wheel_bound = False
        self._options_revision = None
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(36, 10))
        ttk.Label(header, text="添加员工数据", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="录入数据后将保存到“员工数据.xlsx”，可在查询页面查看",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        form_area = ttk.Frame(self, style="Card.TFrame")
        form_area.grid(row=1, column=0, sticky="nsew", padx=36, pady=10)
        form_area.rowconfigure(0, weight=1)
        form_area.columnconfigure(0, weight=1)

        canvas = tk.Canvas(form_area, highlightthickness=0, background="#ffffff")
        vsb = ttk.Scrollbar(form_area, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.canvas = canvas

        self.form = ttk.Frame(canvas, style="Card.TFrame", padding=28)
        self.form.columnconfigure(1, weight=1)
        self._canvas_win = canvas.create_window((0, 0), window=self.form, anchor="nw")
        self.form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind(
            "<Configure>", lambda e: canvas.itemconfigure(self._canvas_win, width=e.width)
        )

        self.name_var = tk.StringVar()
        self.field_widgets = {}
        self.field_combos = {}
        self.rate_vars = {}
        self.name_combo = None

        tip = ttk.Label(
            self,
            text="提示：查询选项内容统一在“新内容添加”维护；达成率、合格率由数量自动计算并标红。",
            style="Sub.TLabel",
        )
        tip.grid(row=2, column=0, sticky="sw", padx=38, pady=(2, 0))
        self._build_form()

    def _build_form(self):
        for child in self.form.winfo_children():
            child.destroy()
        self.field_widgets = {}
        self.field_combos = {}
        self.rate_vars = {}
        headers = self.app.data.columns()
        all_values = self.app.data.all_column_values()
        option_values = {
            field: self.app.options.values(field) for field in OPTION_FIELDS
        }
        row = 0

        ttk.Label(self.form, text="员工姓名：").grid(
            row=row, column=0, sticky="e", padx=(0, 12), pady=8
        )
        self.name_combo = FilterCombobox(
            self.form,
            textvariable=self.name_var,
            font=(FONT_FAMILY, 12),
            values=option_values["员工姓名"],
            state="readonly",
        )
        self.name_combo.grid(row=row, column=1, sticky="ew", pady=8)
        row += 1

        ttk.Label(self.form, text="日期：").grid(
            row=row, column=0, sticky="e", padx=(0, 12), pady=8
        )
        date_box = ttk.Frame(self.form)
        date_box.grid(row=row, column=1, sticky="w", pady=8)
        self.date_entry = ClearableDateEntry(
            date_box, date_pattern="yyyy-mm-dd", locale="zh_CN", width=16
        )
        self.date_entry.pack(side="left")
        ttk.Button(
            date_box,
            text="今天",
            width=6,
            command=lambda: self.date_entry.set_date(datetime.date.today()),
        ).pack(side="left", padx=(8, 0))
        row += 1

        for header in headers:
            if header in (REQUIRED_NAME, REQUIRED_DATE):
                continue
            if header in RATE_COLUMNS:
                var = tk.StringVar(value="—")
                ttk.Label(
                    self.form,
                    textvariable=var,
                    foreground="#d32f2f",
                    font=(FONT_FAMILY, 12, "bold"),
                ).grid(row=row, column=1, sticky="w", pady=8)
                self.rate_vars[header] = var
                row += 1
                continue
            if header in YES_NO_COLUMNS:
                ttk.Label(self.form, text=f"{header}：").grid(
                    row=row, column=0, sticky="e", padx=(0, 12), pady=8
                )
                var = tk.StringVar(value="否")
                combo = FilterCombobox(
                    self.form,
                    textvariable=var,
                    font=(FONT_FAMILY, 12),
                    values=("否", "是"),
                    state="readonly",
                    allow_input=False,
                    width=8,
                )
                combo._no_clear = True
                combo.grid(row=row, column=1, sticky="w", pady=8)
                self.field_widgets[header] = var
                self.field_combos[header] = combo
                row += 1
                continue
            ttk.Label(self.form, text=f"{header}：").grid(
                row=row, column=0, sticky="e", padx=(0, 12), pady=8
            )
            var = tk.StringVar()
            combo = FilterCombobox(
                self.form,
                textvariable=var,
                font=(FONT_FAMILY, 12),
                state="readonly" if header in OPTION_FILES else "normal",
                values=(
                    option_values.get(header, [])
                    if header in OPTION_FIELDS
                    else all_values.get(header, [])
                ),
            )
            combo.grid(row=row, column=1, sticky="ew", pady=8)
            self.field_widgets[header] = var
            self.field_combos[header] = combo
            row += 1

        for key in ("当日计划产量", "实际产量", "合格品数量", "不良品数量"):
            var = self.field_widgets.get(key)
            if var is not None:
                var.trace_add("write", lambda *a: self._update_rates())
        self._update_rates()

        buttons = ttk.Frame(self.form)
        buttons.grid(row=row, column=1, sticky="w", pady=(16, 4))
        ttk.Button(buttons, text="保 存", style="Primary.TButton", command=self.save).pack(
            side="left", padx=(0, 12)
        )
        ttk.Button(buttons, text="清 空", command=self.clear_form).pack(side="left", padx=(0, 12))
        ttk.Button(buttons, text="返回查询", command=lambda: self.app.show("query")).pack(side="left")
        bind_clipboard_shortcuts(self.form)
        bind_combobox_behavior(self.form)
        self._options_revision = self.app.options.revision
        self._data_signature = self.app.data.file_signature()

    def on_show(self):
        self._refresh_options()
        if not self._page_wheel_bound:
            self._bind_page_wheel()
            self._page_wheel_bound = True
        self.date_entry.set_date(datetime.date.today())
        self.name_combo.focus_set()

    def _bind_page_wheel(self):
        for widget in _walk_widgets(self):
            for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                widget.bind(sequence, self._on_page_wheel)
        wheel_command = self.register(self._on_popdown_wheel)
        for combo in (
            widget
            for widget in _walk_widgets(self.form)
            if widget.winfo_class() == "TCombobox"
        ):
            try:
                popdown = combo.tk.call(
                    "ttk::combobox::PopdownWindow", combo._w
                )
                listbox = str(popdown) + ".f.l"
                combo.tk.call(
                    "bind",
                    listbox,
                    "<MouseWheel>",
                    f"{wheel_command} %W %D %X %Y",
                )
                combo.tk.call(
                    "bind",
                    listbox,
                    "<Button-4>",
                    f"{wheel_command} %W -120 %X %Y",
                )
                combo.tk.call(
                    "bind",
                    listbox,
                    "<Button-5>",
                    f"{wheel_command} %W 120 %X %Y",
                )
            except tk.TclError:
                pass

    @staticmethod
    def _wheel_steps(event):
        if getattr(event, "num", None) == 4:
            return -1
        if getattr(event, "num", None) == 5:
            return 1
        delta = getattr(event, "delta", 0)
        if not delta:
            return 0
        return -1 if delta > 0 else 1

    @staticmethod
    def _delta_steps(delta):
        try:
            value = int(delta)
        except (TypeError, ValueError):
            return 0
        if not value:
            return 0
        return -1 if value > 0 else 1

    def _on_popdown_wheel(self, widget_path, delta, x_root, y_root):
        steps = self._delta_steps(delta)
        if not steps:
            return "break"
        try:
            popdown = str(self.tk.call("winfo", "toplevel", widget_path))
            try:
                x = int(x_root)
                y = int(y_root)
            except (TypeError, ValueError):
                return "break"
            pop_x = int(self.tk.call("winfo", "rootx", popdown))
            pop_y = int(self.tk.call("winfo", "rooty", popdown))
            pop_w = int(self.tk.call("winfo", "width", popdown))
            pop_h = int(self.tk.call("winfo", "height", popdown))
            if pop_x <= x < pop_x + pop_w and pop_y <= y < pop_y + pop_h:
                self.tk.call(widget_path, "yview", "scroll", steps, "units")
                return "break"

            self._close_page_dropdowns()
            page_x = int(self.tk.call("winfo", "rootx", self))
            page_y = int(self.tk.call("winfo", "rooty", self))
            page_w = int(self.tk.call("winfo", "width", self))
            page_h = int(self.tk.call("winfo", "height", self))
            if page_x <= x < page_x + page_w and page_y <= y < page_y + page_h:
                self.canvas.yview_scroll(steps, "units")
        except tk.TclError:
            pass
        return "break"

    def _find_combobox(self, widget):
        while widget is not None:
            if isinstance(widget, ttk.Combobox):
                return widget
            widget = getattr(widget, "master", None)
        return None

    @staticmethod
    def _popdown_is_visible(combo):
        try:
            popdown = combo.tk.call(
                "ttk::combobox::PopdownWindow", combo._w
            )
            return bool(int(combo.tk.call("winfo", "ismapped", popdown)))
        except (tk.TclError, ValueError):
            return False

    def _close_page_dropdowns(self):
        for widget in _walk_widgets(self.form):
            if widget.winfo_class() != "TCombobox":
                continue
            try:
                widget.tk.call("ttk::combobox::Unpost", widget._w)
            except tk.TclError:
                pass

    def _on_page_wheel(self, event):
        combo = self._find_combobox(event.widget)
        if combo is not None and self._popdown_is_visible(combo):
            try:
                popdown = combo.tk.call(
                    "ttk::combobox::PopdownWindow", combo._w
                )
                listbox = str(popdown) + ".f.l"
                steps = self._wheel_steps(event)
                if steps:
                    combo.tk.call(listbox, "yview", "scroll", steps, "units")
            except tk.TclError:
                pass
        else:
            self._close_page_dropdowns()
            steps = self._wheel_steps(event)
            if steps:
                self.canvas.yview_scroll(steps, "units")
        return "break"

    def _refresh_options(self):
        data_signature = self.app.data.file_signature()
        if (
            self._options_revision == self.app.options.revision
            and self._data_signature == data_signature
        ):
            return
        all_values = self.app.data.all_column_values()
        if self.name_combo is not None:
            set_combo_values(
                self.name_combo,
                self.app.options.values("员工姓名"),
            )
        for header, combo in self.field_combos.items():
            if header in YES_NO_COLUMNS:
                continue
            if header in OPTION_FIELDS:
                set_combo_values(
                    combo, self.app.options.values(header)
                )
            else:
                set_combo_values(combo, all_values.get(header, []))
        self._options_revision = self.app.options.revision
        self._data_signature = data_signature

    def clear_form(self, keep_name=False, keep_date=False):
        if not keep_date:
            self.date_entry.set_date(datetime.date.today())
        for header, var in self.field_widgets.items():
            var.set("否" if header in YES_NO_COLUMNS else "")
        if not keep_name:
            self.name_var.set("")
            self.name_combo.focus_set()

    def _update_rates(self):
        if "达成率" not in self.rate_vars or "合格率" not in self.rate_vars:
            return
        def val(key):
            var = self.field_widgets.get(key)
            return var.get() if var is not None else None

        rate, quality = calc_rates(
            val("当日计划产量"), val("实际产量"), val("合格品数量"), val("不良品数量")
        )
        self.rate_vars["达成率"].set(rate)
        self.rate_vars["合格率"].set(quality)

    def save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入员工姓名。")
            return
        picked_date = get_date_or_none(self.date_entry)
        if not picked_date:
            messagebox.showwarning("提示", "请选择日期。")
            return
        date_norm = picked_date.isoformat()
        record = {REQUIRED_NAME: name, REQUIRED_DATE: date_norm}
        for header, var in self.field_widgets.items():
            record[header] = var.get().strip()
        for header, var in self.rate_vars.items():
            value = var.get().strip()
            record[header] = "" if value == "—" else value
        if not messagebox.askyesno(
            "确认保存",
            f"是否保存这条员工数据？\n\n员工姓名：{name}\n日期：{date_norm}",
        ):
            return
        try:
            self.app.data.add_row(record)
        except AppError as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self._refresh_options()
        self.clear_form(keep_date=True)


class OptionEditorDialog(tk.Toplevel):
    """管理一个字段的全部内部下拉选项。"""

    def __init__(self, app, field, on_changed):
        super().__init__(app)
        self.app = app
        self.field = field
        self.on_changed = on_changed
        self.title(f"管理选项 - {field}")
        self.geometry("1000x700")
        self.minsize(820, 600)
        self.configure(bg="#eef8f4")
        self.transient(app)
        self.protocol("WM_DELETE_WINDOW", self._close)

        header = ttk.Frame(self)
        header.pack(fill="x", padx=24, pady=(20, 10))
        ttk.Label(header, text=f"{field}选项管理", style="Title.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            header,
            text="新增或删除后，查询页和添加数据页会自动使用最新选项",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        search_bar = ttk.Frame(self, style="Card.TFrame", padding=12)
        search_bar.pack(fill="x", padx=24, pady=(0, 8))
        search_bar.columnconfigure(1, weight=1)
        ttk.Label(search_bar, text="搜索：").grid(row=0, column=0, padx=(0, 8))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(
            search_bar, textvariable=self.search_var
        )
        search_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(
            search_bar, text="刷新", command=self._refresh_tree
        ).grid(row=0, column=2, padx=(8, 0))
        self.search_var.trace_add("write", lambda *args: self._refresh_tree())

        tree_box = ttk.Frame(self, style="Card.TFrame", padding=10)
        tree_box.rowconfigure(0, weight=1)
        tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            tree_box,
            columns=("value", "status", "usage"),
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("value", text="选项内容")
        self.tree.heading("status", text="状态")
        self.tree.heading("usage", text="历史使用条数")
        self.tree.column("value", width=380, anchor="w")
        self.tree.column("status", width=90, anchor="center")
        self.tree.column("usage", width=120, anchor="center")
        scroll = ttk.Scrollbar(
            tree_box, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.tag_configure("inactive", foreground="#8b9aa0")
        self.tree.tag_configure("active", foreground="#27434b")

        add_bar = ttk.Frame(self, style="Card.TFrame", padding=12)
        add_bar.columnconfigure(1, weight=1)
        ttk.Label(add_bar, text="新内容：").grid(row=0, column=0, padx=(0, 8))
        self.new_value_var = tk.StringVar()
        self.new_entry = ttk.Entry(
            add_bar, textvariable=self.new_value_var
        )
        self.new_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(
            add_bar,
            text="添加",
            style="Primary.TButton",
            command=self._add,
        ).grid(row=0, column=2, padx=(8, 0))

        actions = ttk.Frame(self)
        ttk.Button(
            actions, text="删除选中", command=lambda: self._set_active(False)
        ).pack(side="left")
        ttk.Button(
            actions,
            text="彻底删除",
            style="Danger.TButton",
            command=self._delete_permanently,
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            actions, text="恢复选中", command=lambda: self._set_active(True)
        ).pack(side="left", padx=(10, 0))
        ttk.Button(actions, text="关闭", command=self._close).pack(side="right")
        actions.pack(side="bottom", fill="x", padx=24, pady=(4, 14))
        add_bar.pack(side="bottom", fill="x", padx=24, pady=6)
        tree_box.pack(fill="both", expand=True, padx=24, pady=8)

        self._items = {}
        self.new_entry.bind("<Return>", lambda event: self._add())
        self.bind("<Escape>", lambda event: self._close())
        bind_clipboard_shortcuts(self)
        self._refresh_tree()
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1000, max(820, screen_w - 80))
        height = min(700, max(600, screen_h - 100))
        x = app.winfo_rootx() + max(0, (app.winfo_width() - width) // 2)
        y = app.winfo_rooty() + max(0, (app.winfo_height() - height) // 2)
        x = min(max(x, 0), max(0, screen_w - width))
        y = min(max(y, 0), max(0, screen_h - height))
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.grab_set()
        self.after(50, self.new_entry.focus_set)

    def _usage_counts(self):
        counts = {}
        for row in self.app.data.rows():
            value = str(row.get(self.field, "") or "").strip()
            if value:
                counts[value] = counts.get(value, 0) + 1
        return counts

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self._items = {}
        needle = self.search_var.get().strip().casefold()
        counts = self._usage_counts()
        for index, entry in enumerate(
            self.app.options.entries(self.field, include_inactive=True)
        ):
            if needle and needle not in entry["value"].casefold():
                continue
            iid = f"item_{index}"
            status = "启用" if entry["active"] else "已删除"
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    entry["value"],
                    status,
                    counts.get(entry["value"], 0),
                ),
                tags=("active" if entry["active"] else "inactive",),
            )
            self._items[iid] = entry

    def _selected_entries(self):
        return [
            self._items[iid]
            for iid in self.tree.selection()
            if iid in self._items
        ]

    def _add(self):
        value = self.new_value_var.get()
        try:
            self.app.options.add(self.field, value)
        except AppError as exc:
            messagebox.showwarning("无法添加", str(exc), parent=self)
            return
        self.new_value_var.set("")
        self._refresh_tree()
        self.on_changed()
        self.new_entry.focus_set()

    def _set_active(self, active):
        entries = self._selected_entries()
        if not entries:
            messagebox.showinfo("提示", "请先选择内容。", parent=self)
            return
        entries = [
            entry for entry in entries if bool(entry["active"]) != bool(active)
        ]
        if not entries:
            return
        if not active:
            counts = self._usage_counts()
            used = [
                (entry["value"], counts.get(entry["value"], 0))
                for entry in entries
                if counts.get(entry["value"], 0)
            ]
            detail = ""
            if used:
                detail = "\n\n以下内容已有历史记录：\n" + "\n".join(
                    f"{value}：{count} 条" for value, count in used
                )
            if not messagebox.askyesno(
                "确认删除",
                "删除后只会从下拉选项中移除，不会删除员工历史记录。"
                + detail
                + "\n\n是否继续？",
                parent=self,
            ):
                return
        for entry in entries:
            self.app.options.set_active(
                self.field, entry["value"], active
            )
        self._refresh_tree()
        self.on_changed()

    def _delete_permanently(self):
        entries = self._selected_entries()
        if not entries:
            messagebox.showinfo("提示", "请先选择内容。", parent=self)
            return
        counts = self._usage_counts()
        used = [
            (entry["value"], counts.get(entry["value"], 0))
            for entry in entries
            if counts.get(entry["value"], 0)
        ]
        usage_text = ""
        if used:
            usage_text = "\n\n以下内容已有历史记录：\n" + "\n".join(
                f"{value}：{count} 条" for value, count in used
            )
        if not messagebox.askyesno(
            "彻底删除",
            f"将彻底删除选中的 {len(entries)} 项内容。"
            "\n删除后无法恢复。"
            "\n员工历史记录不会被删除。"
            + usage_text
            + "\n\n确定继续吗？",
            parent=self,
        ):
            return
        self.app.options.delete_permanently(
            self.field, [entry["value"] for entry in entries]
        )
        self._refresh_tree()
        self.on_changed()

    def _close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()


class OptionManagerPage(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.buttons = {}
        header = ttk.Frame(self)
        header.pack(fill="x", padx=36, pady=(28, 10))
        ttk.Label(header, text="新内容添加", style="Title.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            header,
            text="点击选项类型按钮，可新增、删除或恢复下拉内容",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))
        note = ttk.Label(
            self,
            text="此页维护程序内部选项，不再读取外部 TXT 文件。",
            style="Sub.TLabel",
        )
        note.pack(anchor="w", padx=36, pady=(0, 12))
        grid = ttk.Frame(self, style="Card.TFrame", padding=24)
        grid.pack(fill="both", expand=True, padx=36, pady=(0, 28))
        for index, field in enumerate(OPTION_FIELDS):
            grid.rowconfigure(index // 4, weight=1)
            grid.columnconfigure(index % 4, weight=1)
            button = ttk.Button(
                grid,
                text=field,
                style="Option.TButton",
                command=lambda value=field: self._open(value),
            )
            button.grid(
                row=index // 4,
                column=index % 4,
                sticky="nsew",
                padx=8,
                pady=8,
            )
            self.buttons[field] = button

    def _open(self, field):
        OptionEditorDialog(self.app, field, self._refresh_counts)

    def _refresh_counts(self):
        for field, button in self.buttons.items():
            button.configure(
                text=f"{field}\n{self.app.options.count(field)} 项"
            )

    def on_show(self):
        self._refresh_counts()


class HistoryPage(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        header = ttk.Frame(self)
        header.pack(fill="x", padx=36, pady=(28, 10))
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header, text="历史记录", style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="按录入顺序显示最近 50 条记录",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        header_actions = ttk.Frame(header)
        header_actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(
            header_actions,
            text="删除选中",
            style="Danger.TButton",
            command=self._delete_selected,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            header_actions,
            text="刷新",
            command=lambda: self.on_show(force=True),
        ).pack(side="left")

        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.pack(fill="both", expand=True, padx=36, pady=(0, 20))
        card.rowconfigure(0, weight=1)
        card.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            card, show="headings", selectmode="extended"
        )
        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.status_var = tk.StringVar()
        self._loaded_signature = None
        self._rows_by_iid = {}
        self.tree.bind("<Delete>", lambda event: self._delete_selected())
        self.tree.bind(
            "<<TreeviewSelect>>",
            lambda event: self._update_status(),
        )
        ttk.Label(
            self, textvariable=self.status_var, style="Sub.TLabel"
        ).pack(anchor="e", padx=38, pady=(0, 10))

    def on_show(self, force=False):
        signature = self.app.data.file_signature()
        if not force and self._loaded_signature == signature:
            return
        rows = sorted(
            self.app.data.rows(),
            key=lambda row: int(row.get("_row", 0)),
            reverse=True,
        )[:50]
        columns = self.app.data.columns()
        self.tree.delete(*self.tree.get_children())
        self._rows_by_iid = {}
        self.tree["columns"] = columns
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(
                column,
                width=120,
                minwidth=70,
                anchor="center" if column == REQUIRED_DATE else "w",
                stretch=False,
            )
        for index, row in enumerate(rows):
            iid = f"history_{index}"
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=[row.get(column, "") for column in columns],
                tags=("odd" if index % 2 else "even",),
            )
            self._rows_by_iid[iid] = row
        self.tree.tag_configure("odd", background="#f2f9f7")
        self.tree.tag_configure("even", background="#fbfefd")
        self._update_status()
        self._loaded_signature = signature

    def _update_status(self):
        selected = len(self.tree.selection())
        self.status_var.set(
            f"显示最近 {len(self._rows_by_iid)} 条记录"
            f"（最多 50 条），已选 {selected} 条"
        )

    def _delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo(
                "提示", "请先选择要删除的历史记录。", parent=self
            )
            return
        rows = [
            self._rows_by_iid[iid]
            for iid in selected
            if iid in self._rows_by_iid
        ]
        indexes = [row["_row"] for row in rows if "_row" in row]
        if not indexes:
            return
        if not messagebox.askyesno(
            "确认删除",
            f"将删除选中的 {len(indexes)} 条历史记录。"
            "\n删除后无法恢复，是否继续？",
            parent=self,
        ):
            return
        try:
            self.app.data.delete_rows_by_indexes(indexes)
        except AppError as exc:
            messagebox.showerror("删除失败", str(exc), parent=self)
            return
        self.on_show(force=True)


class TableImportPreviewDialog(tk.Toplevel):
    """预览并确认批量导入 Excel 记录。"""

    def __init__(
        self,
        app,
        parent_page,
        source_path,
        sheet_name,
        valid_records,
        issues,
    ):
        super().__init__(app)
        self.app = app
        self.parent_page = parent_page
        self.source_path = source_path
        self.sheet_name = sheet_name
        self.valid_records = list(valid_records)
        self.issues = list(issues)
        self.preview_items = {}
        self.title("Excel 导入预览")
        self.geometry("1000x650")
        self.minsize(860, 560)
        self.resizable(True, True)
        self.configure(bg="#eef8f4")
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        header = ttk.Frame(self)
        header.pack(fill="x", padx=24, pady=(20, 8))
        title_area = ttk.Frame(header)
        title_area.pack(side="left", fill="x", expand=True)
        ttk.Label(
            title_area, text="Excel 导入预览", style="Title.TLabel"
        ).pack(anchor="w")
        self.source_var = tk.StringVar(
            value=(
                f"来源：{os.path.basename(source_path)}  "
                f"工作表：{sheet_name}"
            )
        )
        ttk.Label(
            title_area,
            textvariable=self.source_var,
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))
        ttk.Label(
            title_area,
            text="双击任意预览行可打开原始 Excel；可多选后删除本次预览行。",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        summary = ttk.Frame(self, style="Card.TFrame", padding=12)
        summary.pack(fill="x", padx=24, pady=8)
        summary.columnconfigure(1, weight=1)
        self.summary_var = tk.StringVar()
        ttk.Label(
            summary,
            textvariable=self.summary_var,
            font=(FONT_FAMILY, 12, "bold"),
            foreground="#2f6f82",
        ).grid(row=0, column=0, sticky="nw")
        self.error_reason_var = tk.StringVar()
        ttk.Label(
            summary,
            textvariable=self.error_reason_var,
            font=(FONT_FAMILY, 11, "bold"),
            foreground="#d32f2f",
            justify="left",
            wraplength=760,
        ).grid(row=0, column=1, sticky="nw", padx=(18, 0))

        box = ttk.Frame(self, style="Card.TFrame", padding=10)
        box.pack(fill="both", expand=True, padx=24, pady=8)
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        columns = ("row",) + IMPORT_COLUMNS
        self.tree = ttk.Treeview(
            box, columns=columns, show="headings", selectmode="extended"
        )
        headings = {
            "row": "Excel行",
        }
        widths = {
            "row": 70,
        }
        for field in IMPORT_COLUMNS:
            headings[field] = field
            widths[field] = RESULT_COLUMN_WIDTHS.get(field, 110)
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(
                column,
                width=widths[column],
                minwidth=60,
                anchor="center",
                stretch=False,
            )
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("valid", background="#f2f9f7")
        self.tree.tag_configure("duplicate", background="#fff6df")
        self.tree.tag_configure("error", background="#fff0f0")
        self._populate_preview()
        self.tree.bind("<Double-1>", self._open_source_workbook)
        self.tree.bind(
            "<Delete>",
            lambda event: (
                self._delete_selected_rows() or "break"
            ),
        )

        actions = ttk.Frame(self)
        actions.pack(fill="x", padx=24, pady=(4, 18))
        ttk.Button(
            actions,
            text="删除选中行",
            command=self._delete_selected_rows,
        ).pack(side="left")
        ttk.Button(
            actions,
            text="刷新预览",
            command=self._refresh_preview,
        ).pack(side="left", padx=(10, 0))
        ttk.Label(
            actions,
            text="双击行可打开原 Excel",
            style="Sub.TLabel",
        ).pack(side="left", padx=(12, 0))
        ttk.Button(
            actions, text="取消", command=self.destroy
        ).pack(side="right")
        self.import_button = ttk.Button(
            actions,
            text="确认导入",
            style="Primary.TButton",
            command=self._commit,
        )
        self.import_button.pack(side="right", padx=(0, 10))
        self._refresh_summary()

        self.update_idletasks()
        width, height = 1000, 650
        x = app.winfo_rootx() + max(0, (app.winfo_width() - width) // 2)
        y = app.winfo_rooty() + max(0, (app.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")
        self.grab_set()
        self.tree.focus_set()

    def _populate_preview(self):
        children = self.tree.get_children()
        if children:
            self.tree.delete(*children)
        self.preview_items.clear()
        preview = [
            (
                record["_source_row"],
                record,
                "valid",
                "valid",
                record,
            )
            for record in self.valid_records
        ]
        preview.extend(
            (
                issue["row"],
                issue["record"],
                "duplicate" if issue["status"] == "重复" else "error",
                "issue",
                issue,
            )
            for issue in self.issues
        )
        for row_number, record, tag, kind, source in sorted(
            preview, key=lambda item: item[0]
        ):
            values = [row_number]
            values.extend(record.get(field, "") for field in IMPORT_COLUMNS)
            item_id = self.tree.insert(
                "",
                "end",
                values=tuple(values),
                tags=(tag,),
            )
            self.preview_items[item_id] = (kind, source)

    def _refresh_preview(self):
        if not os.path.exists(self.source_path):
            messagebox.showerror(
                "无法刷新预览",
                f"原始 Excel 文件已不存在：\n{self.source_path}",
                parent=self,
            )
            return
        try:
            sheet_name, valid_records, issues = (
                self.parent_page._parse_workbook(self.source_path)
            )
        except Exception as exc:
            messagebox.showerror(
                "刷新预览失败",
                str(exc),
                parent=self,
            )
            return
        self.sheet_name = sheet_name
        self.valid_records = list(valid_records)
        self.issues = list(issues)
        self.source_var.set(
            f"来源：{os.path.basename(self.source_path)}  "
            f"工作表：{self.sheet_name}"
        )
        TableImportPage._warn_for_date_issues(self.issues, self)
        self._populate_preview()
        self._refresh_summary()
        self.tree.focus_set()

    def _refresh_summary(self):
        duplicate_count = sum(
            1 for issue in self.issues if issue["status"] == "重复"
        )
        invalid_count = sum(
            1 for issue in self.issues if issue["status"] == "错误"
        )
        self.summary_var.set(
            f"可导入：{len(self.valid_records)} 条    "
            f"重复跳过：{duplicate_count} 条    "
            f"格式错误：{invalid_count} 条"
        )
        error_reasons = [
            f"第 {issue['row']} 行：{issue.get('reason', '')}"
            for issue in sorted(
                (
                    issue
                    for issue in self.issues
                    if issue["status"] == "错误"
                ),
                key=lambda issue: issue["row"],
            )
        ]
        self.error_reason_var.set(
            "错误原因：" + "；".join(error_reasons)
            if error_reasons
            else ""
        )
        self.import_button.configure(
            text=f"确认导入 {len(self.valid_records)} 条"
        )
        if self.valid_records:
            self.import_button.state(["!disabled"])
        else:
            self.import_button.state(["disabled"])

    def _delete_selected_rows(self):
        selected = list(self.tree.selection())
        if not selected:
            messagebox.showinfo(
                "未选择行",
                "请先选中一条或多条预览记录。",
                parent=self,
            )
            return
        if not messagebox.askyesno(
            "删除预览行",
            f"从本次导入预览中移除选中的 {len(selected)} 行？\n"
            "原 Excel 文件不会被修改。",
            parent=self,
        ):
            return
        for item_id in selected:
            kind, source = self.preview_items.pop(item_id, (None, None))
            if kind == "valid":
                self.valid_records[:] = [
                    record
                    for record in self.valid_records
                    if record is not source
                ]
            elif kind == "issue":
                self.issues[:] = [
                    issue for issue in self.issues if issue is not source
                ]
            self.tree.delete(item_id)
        self._refresh_summary()

    def _open_source_workbook(self, event=None):
        if event is not None and not self.tree.identify_row(event.y):
            return
        if not os.path.exists(self.source_path):
            messagebox.showerror(
                "无法打开Excel",
                f"原始 Excel 文件已不存在：\n{self.source_path}",
                parent=self,
            )
            return "break"
        try:
            os.startfile(self.source_path)
        except OSError as exc:
            messagebox.showerror(
                "无法打开Excel",
                f"{exc}\n\n文件位置：\n{self.source_path}",
                parent=self,
            )
        return "break"

    def _commit(self):
        if not self.valid_records:
            return
        backup_dir = os.path.join(base_dir(), "查询结果", "自动备份")
        os.makedirs(backup_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        import shutil

        data_backup = os.path.join(
            backup_dir, f"员工数据_导入前_{stamp}.xlsx"
        )
        option_backup = os.path.join(
            backup_dir, f"选项内容_导入前_{stamp}.dat"
        )
        try:
            shutil.copy2(self.app.data.path, data_backup)
            if os.path.exists(self.app.options.path):
                shutil.copy2(self.app.options.path, option_backup)
            imported = self.app.data.add_rows_batch(self.valid_records)
        except Exception as exc:
            messagebox.showerror(
                "导入失败",
                f"{exc}\n\n导入前备份保存在：\n{backup_dir}",
                parent=self,
            )
            return
        self.parent_page.status_var.set(
            f"已导入 {imported} 条记录，备份：{os.path.basename(data_backup)}"
        )
        self.grab_release()
        self.destroy()
        messagebox.showinfo(
            "导入完成",
            f"成功导入 {imported} 条有效记录。\n"
            f"非内部选项、格式错误和重复记录均未导入。\n\n"
            f"导入前备份目录：\n{backup_dir}",
            parent=self.app,
        )


class TableImportPage(ttk.Frame):
    """生成 Excel 填写模板并预览导入。"""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        header = ttk.Frame(self)
        header.pack(fill="x", padx=36, pady=(28, 10))
        ttk.Label(
            header, text="表格录入", style="Title.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="生成Excel填写模板，或直接导入别人填写完成的数据",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        cards = ttk.Frame(self)
        cards.pack(fill="both", expand=True, padx=36, pady=(10, 20))
        cards.columnconfigure(0, weight=1, uniform="import_cards")
        cards.columnconfigure(1, weight=1, uniform="import_cards")

        template_card = ttk.Frame(
            cards, style="BlueSection.TFrame", padding=(24, 22)
        )
        template_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(
            template_card,
            text="第一步：生成填写表格",
            style="BlueTitle.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            template_card,
            text=(
                "生成包含员工姓名、日期、生产数据等字段的 Excel。\n"
                "文件会保存到程序目录下的“表格模板”文件夹。"
            ),
            style="CardSub.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(12, 18))
        template_actions = ttk.Frame(template_card, style="BlueSection.TFrame")
        template_actions.pack(anchor="w")
        ttk.Button(
            template_actions,
            text="生成Excel模板",
            style="Primary.TButton",
            command=self.generate_template,
        ).pack(side="left")

        import_card = ttk.Frame(
            cards, style="GreenSection.TFrame", padding=(24, 22)
        )
        import_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(
            import_card,
            text="第二步：导入填写后的表格",
            style="GreenTitle.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            import_card,
            text=(
                "选择别人填写后的 Excel，先检查内部选项、必填项、\n"
                "日期、数值和重复记录，确认后才会批量写入程序。"
            ),
            style="CardSub.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(12, 18))
        ttk.Button(
            import_card,
            text="导入Excel表格",
            style="Primary.TButton",
            command=self.preview_import,
        ).pack(anchor="w")

        self.status_var = tk.StringVar(
            value="尚未导入；填写完成后点击“导入Excel表格”"
        )
        ttk.Label(
            self, textvariable=self.status_var, style="Sub.TLabel"
        ).pack(anchor="e", padx=38, pady=(0, 12))

    def generate_template(self):
        out_dir = os.path.join(base_dir(), TEMPLATE_DIR)
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(
            out_dir,
            "员工产能填写模板_{:%Y%m%d_%H%M%S}.xlsx".format(
                datetime.datetime.now()
            ),
        )
        wb = Workbook()
        ws = wb.active
        ws.title = "数据录入"
        ws.append(IMPORT_COLUMNS)
        style_header_row(ws, len(IMPORT_COLUMNS))
        for index, column in enumerate(IMPORT_COLUMNS, 1):
            ws.column_dimensions[get_column_letter(index)].width = (
                18 if column == "备注" else 15
            )
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(len(IMPORT_COLUMNS))}1"
        )
        if DataValidation is not None:
            date_validation = DataValidation(
                type="date",
                operator="between",
                formula1="DATE(2000,1,1)",
                formula2="DATE(2099,12,31)",
                allow_blank=True,
            )
            date_validation.promptTitle = "填写日期"
            date_validation.prompt = (
                "请填写有效日期，格式为 YYYY-MM-DD，例如 2026-11-06。"
            )
            date_validation.errorTitle = "日期格式不正确"
            date_validation.error = (
                "请填写有效日期，格式为 YYYY-MM-DD，例如 2026-11-06。"
            )
            date_validation.showInputMessage = True
            date_validation.showErrorMessage = True
            ws.add_data_validation(date_validation)
            date_validation.add("B2:B201")
            yes_no = DataValidation(
                type="list", formula1='"否,是"', allow_blank=True
            )
            ws.add_data_validation(yes_no)
            ws["D2"].value = "否"
            ws["E2"].value = "否"
            yes_no.add("D2:E201")
            for column, field in enumerate(IMPORT_COLUMNS, 1):
                if field not in OPTION_FIELDS:
                    continue
                values = self.app.options.values(field)
                if not values:
                    continue
                source_column = get_column_letter(
                    OPTION_FIELDS.index(field) + 1
                )
                last_row = len(values) + 1
                field_validation = DataValidation(
                    type="list",
                    formula1=(
                        f"'可选项'!${source_column}$2:"
                        f"${source_column}${last_row}"
                    ),
                    allow_blank=True,
                )
                field_validation.errorTitle = "内容不在程序内部选项中"
                field_validation.error = (
                    "请从下拉列表选择；新增内容请先在程序的"
                    "“新内容添加”中加入。"
                )
                field_validation.showErrorMessage = True
                ws.add_data_validation(field_validation)
                target_column = get_column_letter(column)
                field_validation.add(f"{target_column}2:{target_column}201")
        for row in range(2, 202):
            ws.cell(row=row, column=2).number_format = "yyyy-mm-dd"
            for column in range(12, 17):
                ws.cell(row=row, column=column).number_format = "0.###"

        options_sheet = wb.create_sheet("可选项")
        for column, field in enumerate(OPTION_FIELDS, 1):
            cell = options_sheet.cell(1, column, field)
            cell.fill = PatternFill("solid", fgColor="5A9CAF")
            cell.font = Font(
                color="FFFFFF", bold=True, name="微软雅黑", size=11
            )
            options_sheet.column_dimensions[
                get_column_letter(column)
            ].width = 20
            for row, value in enumerate(
                self.app.options.values(field), 2
            ):
                options_sheet.cell(row, column, value)
        options_sheet.freeze_panes = "A2"

        info = wb.create_sheet("填写说明")
        tips = [
            ["员工产能报告填写说明"],
            ["1. 在“数据录入”工作表中逐行填写，每行一条员工记录。"],
            ["2. 员工姓名和日期为必填项。"],
            ["3. 日期必须使用 YYYY-MM-DD，例如 2026-11-06。"],
            [
                "4. 员工姓名、工号、班组、机种、物料、客户、工序、模具号"
                "只能选择“可选项”工作表中已有的内部内容。"
            ],
            ["5. 上述下拉字段可以留空；填写了内部选项以外的内容会整行跳过。"],
            ["6. 是否新人、是否支援人员可填“是”或“否”，留空时自动按“否”保存。"],
            ["7. 产量和工时可以留空；填写时必须是数字，否则整行跳过。"],
            ["8. 姓名为空、日期为空或日期格式错误的行会跳过。"],
            ["9. 达成率、合格率无需填写，导入后由程序自动计算。"],
            ["10. 不要修改表头名称，不要在表头前插入标题行。"],
        ]
        for row in tips:
            info.append(row)
        info["A1"].font = Font(
            bold=True, size=14, color="2F6F82", name="微软雅黑"
        )
        info.column_dimensions["A"].width = 90
        wb.save(path)
        self.status_var.set(f"模板已保存：{os.path.basename(path)}")
        if messagebox.askyesno(
            "模板已生成",
            f"Excel填写模板已保存到：\n{path}\n\n"
            "填写完成后，请返回本页点击“导入Excel表格”。\n\n"
            "是否现在打开模板？",
            parent=self,
        ):
            try:
                os.startfile(path)
            except OSError:
                pass

    @staticmethod
    def _header_key(value):
        return re.sub(r"\s+", "", str(value or ""))

    def _parse_workbook(self, path):
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            return self._parse_loaded_workbook(wb)
        finally:
            wb.close()

    def _parse_loaded_workbook(self, wb):
        selected = None
        for sheet in wb.worksheets:
            rows = sheet.iter_rows(
                min_row=1, max_row=min(15, sheet.max_row), values_only=True
            )
            for row_number, values in enumerate(rows, 1):
                mapped = {}
                for index, value in enumerate(values):
                    key = self._header_key(value)
                    canonical = IMPORT_HEADER_ALIASES.get(key)
                    if canonical and canonical not in mapped.values():
                        mapped[index] = canonical
                if (
                    "员工姓名" in mapped.values()
                    and "日期" in mapped.values()
                ):
                    selected = (sheet, row_number, mapped)
                    break
            if selected:
                break
        if selected is None:
            raise AppError(
                "未找到有效表头，请确保Excel包含“员工姓名”和“日期”列。"
            )
        sheet, header_row, mapped = selected
        existing_rows = self.app.data.rows()
        option_maps = {}
        for field in OPTION_FIELDS:
            mapping = {}
            for value in self.app.options.values(field):
                canonical = str(value or "").strip()
                if canonical:
                    mapping.setdefault(canonical.casefold(), canonical)
            option_maps[field] = mapping

        def duplicate_key(record):
            return tuple(
                self._normalize_key(record.get(field))
                for field in (
                    "员工姓名",
                    "日期",
                    "工号",
                    "物料代码",
                    "所属产线/工序",
                    "模具号",
                    "当日计划产量",
                    "实际产量",
                    "合格品数量",
                    "不良品数量",
                )
            )

        existing_keys = {duplicate_key(row) for row in existing_rows}
        file_keys = set()
        valid = []
        issues = []
        numeric_fields = (
            "当日计划产量",
            "实际产量",
            "合格品数量",
            "不良品数量",
            "工时",
        )
        supported_headers = set(IMPORT_COLUMNS)
        for row_number, values in enumerate(
            sheet.iter_rows(
                min_row=header_row + 1,
                values_only=True,
            ),
            header_row + 1,
        ):
            if not any(
                value is not None and str(value).strip()
                for value in values
            ):
                continue
            record = {"_source_row": row_number}
            for index, header in mapped.items():
                record[header] = (
                    values[index] if index < len(values) else None
                )
            for header in supported_headers:
                record.setdefault(header, "")
            for field in supported_headers:
                value = record.get(field)
                if value is None:
                    record[field] = ""
                elif isinstance(value, str):
                    record[field] = value.strip()
            name = str(record.get("员工姓名", "") or "").strip()
            date_value = normalize_import_date(record.get("日期"))
            reasons = []
            if not name:
                reasons.append("员工姓名为空")
            if not date_value:
                reasons.append("日期为空或格式错误，应为YYYY-MM-DD")
            for field in numeric_fields:
                value = record.get(field)
                if (
                    value is not None
                    and str(value).strip()
                    and to_num(value) is None
                ):
                    reasons.append(f"{field}不是数字")
            for field in YES_NO_COLUMNS:
                value = str(record.get(field, "") or "").strip()
                if not value:
                    record[field] = "否"
                elif value not in ("是", "否"):
                    reasons.append(f"{field}只能填是或否")
            for field in OPTION_FIELDS:
                value = str(record.get(field, "") or "").strip()
                if not value:
                    record[field] = ""
                    continue
                canonical = option_maps[field].get(value.casefold())
                if canonical is None:
                    reasons.append(
                        f"{field}“{value}”不在程序内部选项中"
                    )
                else:
                    record[field] = canonical
            if reasons:
                issues.append(
                    {
                        "row": row_number,
                        "record": record,
                        "status": "错误",
                        "reason": "；".join(reasons),
                    }
                )
                continue
            record["员工姓名"] = name
            record["日期"] = date_value
            key = duplicate_key(record)
            if key in existing_keys or key in file_keys:
                issues.append(
                    {
                        "row": row_number,
                        "record": record,
                        "status": "重复",
                        "reason": "与已有记录或文件内其他行重复",
                    }
                )
                continue
            file_keys.add(key)
            valid.append(record)
        return sheet.title, valid, issues

    @staticmethod
    def _normalize_key(value):
        date_value = normalize_date(value)
        if date_value:
            return date_value
        number = to_num(value)
        if number is not None:
            return f"{number:.8f}"
        return str(value or "").strip().casefold()

    @staticmethod
    def _warn_for_date_issues(issues, parent=None):
        rows = sorted(
            {
                int(issue["row"])
                for issue in issues
                if "日期" in str(issue.get("reason", ""))
            }
        )
        if not rows:
            return False
        shown = "、".join(f"第 {row} 行" for row in rows[:12])
        if len(rows) > 12:
            shown += f" 等 {len(rows)} 行"
        messagebox.showwarning(
            "日期格式警告",
            f"以下行的日期为空或格式不正确：\n{shown}\n\n"
            "正确格式示例：2026-11-06\n"
            "这些行不会被导入，请在 Excel 中修改后点击“刷新预览”。",
            parent=parent,
        )
        return True

    def preview_import(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="选择已填写的员工产能Excel",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return
        try:
            sheet_name, valid, issues = self._parse_workbook(path)
        except Exception as exc:
            messagebox.showerror(
                "无法读取Excel", str(exc), parent=self
            )
            return
        TableImportPage._warn_for_date_issues(issues, self)
        TableImportPreviewDialog(
            self.app,
            self,
            path,
            sheet_name,
            valid,
            issues,
        )


class ClearDataPage(ttk.Frame):
    """清除数据选项卡：查询记录详情、勾选删除或清除全部。"""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.parent_page = app.pages["query"]

        ttk.Label(self, text="清除数据", style="Title.TLabel").pack(
            anchor="w", padx=36, pady=(20, 8)
        )

        top = ttk.Frame(self, style="Card.TFrame", padding=12)
        top.pack(fill="x", padx=24, pady=(20, 8))
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        ttk.Label(top, text="员工姓名：").grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=6
        )
        self.name_var = tk.StringVar()
        self.name_combo = FilterCombobox(
            top,
            textvariable=self.name_var,
            values=app.options.values("员工姓名"),
            state="readonly",
            width=16,
        )
        self.name_combo.grid(row=0, column=1, sticky="w", pady=6)

        ttk.Label(top, text="工号：").grid(
            row=0, column=2, sticky="e", padx=(16, 8), pady=6
        )
        self.id_var = tk.StringVar()
        self.id_combo = FilterCombobox(
            top,
            textvariable=self.id_var,
            values=app.options.values("工号"),
            state="readonly",
            width=16,
        )
        self.id_combo.grid(row=0, column=3, sticky="w", pady=6)

        date_filter = ttk.Frame(top, style="Card.TFrame")
        date_filter.grid(
            row=0, column=4, sticky="w", padx=(12, 0), pady=6
        )
        self.date_enabled = tk.BooleanVar(value=False)
        self.date_state_var = tk.StringVar(value="关闭")
        RoundStatusToggle(
            date_filter,
            self.date_enabled,
            command=self._toggle_date,
            bg="#f9fdfc",
        ).pack(side="left")
        ttk.Label(
            date_filter,
            text="日期筛选：",
            font=(FONT_FAMILY, 11, "bold"),
        ).pack(side="left", padx=(4, 0))
        ttk.Label(
            date_filter, textvariable=self.date_state_var
        ).pack(side="left")
        self.start_entry = ClearableDateEntry(
            date_filter,
            date_pattern="yyyy-mm-dd",
            locale="zh_CN",
            state="disabled",
            width=10,
        )
        self.start_entry.pack(side="left", padx=(10, 0))
        ttk.Label(date_filter, text="至").pack(side="left", padx=6)
        self.end_entry = ClearableDateEntry(
            date_filter,
            date_pattern="yyyy-mm-dd",
            locale="zh_CN",
            state="disabled",
            width=10,
        )
        self.end_entry.pack(side="left")

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=24, pady=(0, 8))
        ttk.Button(bar, text="查 询", style="Primary.TButton", command=self._do_query).pack(
            side="left"
        )
        ttk.Button(bar, text="删除选中", command=self._delete_selected).pack(
            side="left", padx=(12, 0)
        )
        ttk.Button(bar, text="一键清除全部", command=self._clear_all).pack(
            side="left", padx=(12, 0)
        )
        self.status_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.status_var, style="Sub.TLabel").pack(side="right")

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(box, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.columns = app.data.columns()
        self.tree["columns"] = self.columns
        for col in self.columns:
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(
                col,
                width=RESULT_COLUMN_WIDTHS.get(col, 120),
                minwidth=RESULT_COLUMN_WIDTHS.get(col, 100),
                stretch=True,
            )
        self.tree.bind("<Delete>", lambda e: self._delete_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_status())

        self.row_items = {}
        bind_clipboard_shortcuts(self)
        bind_combobox_behavior(self)

    def on_show(self):
        set_combo_values(
            self.name_combo, self.app.options.values("员工姓名")
        )
        set_combo_values(self.id_combo, self.app.options.values("工号"))
        self.name_var.set("")
        self.id_var.set("")
        self.date_enabled.set(False)
        self.date_state_var.set("关闭")
        self.start_entry.delete(0, "end")
        self.end_entry.delete(0, "end")
        self.start_entry.config(state="disabled")
        self.end_entry.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.row_items = {}
        self.status_var.set("请选择条件后点击查询")

    def _toggle_date(self):
        enabled = bool(self.date_enabled.get())
        state = "normal" if enabled else "disabled"
        self.start_entry.config(state=state)
        self.end_entry.config(state=state)
        self.date_state_var.set("启用" if enabled else "关闭")
        if not enabled:
            self.start_entry.delete(0, "end")
            self.end_entry.delete(0, "end")

    def _filters(self):
        start = end = None
        if self.date_enabled.get():
            start_date = get_date_or_none(self.start_entry)
            end_date = get_date_or_none(self.end_entry)
            start = start_date.isoformat() if start_date else None
            end = end_date.isoformat() if end_date else None
            if start and end and start > end:
                messagebox.showwarning("提示", "开始日期不能晚于结束日期。", parent=self)
                return None
        return {
            "name": self.name_var.get().strip() or None,
            "emp_id": self.id_var.get().strip() or None,
            "start": start,
            "end": end,
        }

    def _do_query(self):
        filters = self._filters()
        if filters is None:
            return
        rows = self.app.data.rows(
            name=filters["name"],
            emp_id=filters["emp_id"],
            start=filters["start"],
            end=filters["end"],
        )
        self._show_rows(rows)

    def _show_rows(self, rows):
        self.tree.delete(*self.tree.get_children())
        self.row_items = {}
        for i, r in enumerate(rows):
            iid = f"r{i}"
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=[r.get(c, "") for c in self.columns],
                tags=("even" if i % 2 == 0 else "odd",),
            )
            self.row_items[iid] = r
        self.tree.tag_configure("odd", background="#f5f7fb")
        self.tree.tag_configure("even", background="#ffffff")
        self._update_status()

    def _update_status(self):
        selected = len(self.tree.selection())
        self.status_var.set(f"共 {len(self.row_items)} 条记录，已选 {selected} 条")

    def _delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先用鼠标勾选要删除的记录。", parent=self)
            return
        indexes = [self.row_items[iid]["_row"] for iid in selected]
        if not messagebox.askyesno(
            "确认删除",
            f"将删除选中的 {len(indexes)} 条记录，删除后无法恢复。\n确定继续吗？",
            parent=self,
        ):
            return
        try:
            self.app.data.delete_rows_by_indexes(indexes)
        except AppError as exc:
            messagebox.showerror("删除失败", str(exc), parent=self)
            return
        self.parent_page.refresh_data()
        self._do_query()

    def _clear_all(self):
        count = self.app.data.delete_rows(preview=True)
        if count == 0:
            messagebox.showinfo("提示", "当前没有可清除的数据。", parent=self)
            return
        if not messagebox.askyesno(
            "一键清除全部",
            f"将删除数据表中的全部 {count} 条记录，删除后无法恢复。\n确定继续吗？",
            parent=self,
        ):
            return
        try:
            self.app.data.delete_rows()
        except AppError as exc:
            messagebox.showerror("删除失败", str(exc), parent=self)
            return
        self.parent_page.refresh_data()
        self._do_query()


class HelpPage(ttk.Frame):
    """选项卡式操作帮助与推荐使用流程。"""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.sections = []

        header = ttk.Frame(self)
        header.pack(fill="x", padx=36, pady=(24, 10))
        ttk.Label(
            header, text="帮助", style="Title.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="按照下面的流程可以完成日常录入、查询、导入和数据维护。",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        area = ttk.Frame(self)
        area.pack(fill="both", expand=True, padx=30, pady=(4, 14))
        area.rowconfigure(0, weight=1)
        area.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            area,
            background="#eef8f4",
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(
            area, orient="vertical", command=self.canvas.yview
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.content = ttk.Frame(self.canvas)
        self.content.columnconfigure(0, weight=1)
        self._canvas_window = self.canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )
        self.content.bind(
            "<Configure>",
            lambda event: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            ),
        )
        self.canvas.bind(
            "<Configure>",
            lambda event: self.canvas.itemconfigure(
                self._canvas_window, width=event.width
            ),
        )

        self._add_intro()
        self._add_section(
            "一、首次使用：先维护下拉内容",
            [
                "进入“新内容添加”，按需要维护员工姓名、工号、班组名称、"
                "所属机种、物料代码、产品客户、所属产线/工序和模具号。",
                "新增内容后，其他页面的下拉列表会自动同步。"
                "“删除选中”可以恢复，“彻底删除”无法恢复。",
            ],
        )
        self._add_section(
            "二、日常录入：添加生成数据",
            [
                "进入“添加生成数据”，选择日期，并依次选择员工、工号、"
                "班组、机种、物料、客户、工序、模具号等内部选项。",
                "填写计划产量、实际产量、合格品数量、不良品数量和工时。"
                "达成率、合格率会自动计算。",
                "确认保存后，除日期外的填写条件会清空，方便继续录入下一位员工。",
            ],
        )
        self._add_section(
            "三、查询与导出：内容查询",
            [
                "可按员工姓名、工号、班组、机种、物料、客户、工序、"
                "模具号、新人或支援人员组合查询。",
                "“日期筛选”选择开始和结束日期；“日期查询”只查询指定一天，"
                "两个日期功能相互独立。",
                "查询结果会显示明细和汇总，可导出 Excel，也可导出热力图和重点趋势图。",
            ],
        )
        self._add_section(
            "四、批量填写与导入：表格录入",
            [
                "点击“生成Excel模板”，模板保存在“表格模板”文件夹。",
                "日期必须严格填写为 YYYY-MM-DD。下拉字段只能使用程序内部已有内容；"
                "留空允许，但填写外部内容会导致该行不导入。",
                "填写完成后点击“导入Excel表格”。在预览窗口可以刷新、多选删除预览行、"
                "双击行打开原 Excel，最后点击“确认导入”。",
                "错误原因会显示在“格式错误”右侧。重复、格式错误和外部选项均不会导入。",
            ],
        )
        self._add_section(
            "五、历史记录与删除数据",
            [
                "“历史记录”显示最近 50 条录入内容，可刷新或删除选中记录。",
                "“清除数据”可按员工姓名、工号和日期范围筛选，再勾选删除；"
                "删除后无法恢复。",
                "“一键清除全部”会删除所有员工数据，使用前请确认。",
            ],
        )
        self._add_section(
            "六、直接查看员工数据",
            [
                "双击本文件夹中的“员工数据解密工具.exe”，输入密码说明文档中的"
                "解密密码，即可导出普通 Excel。",
                "工具不会修改或删除原员工数据；导出的明文 Excel 使用后应妥善保管。",
            ],
        )
        self._add_section(
            "七、重要注意事项",
            [
                "不要删除“员工数据.xlsx”“选项内容.dat”“程序授权.dat”或“_internal”，"
                "否则会分别造成数据、选项、授权或程序运行异常。",
                "分发给其他电脑时必须发送整个“新人员管理系统”文件夹，"
                "不能只发送 EXE。",
                "解密工具、密码文档和密钥生成器仅供管理员使用，"
                "不要放入给普通使用者的分发包。",
            ],
        )

        for widget in _walk_widgets(self.content):
            for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                widget.bind(sequence, self._on_mousewheel)

    def _add_intro(self):
        card = ttk.Frame(
            self.content, style="Card.TFrame", padding=(22, 16)
        )
        card.pack(fill="x", pady=(0, 9), padx=2)
        ttk.Label(
            card,
            text="推荐操作顺序",
            font=(FONT_FAMILY, 12, "bold"),
            foreground="#2f6f82",
        ).pack(anchor="w")
        ttk.Label(
            card,
            text=(
                "维护内部选项 → 手工添加或批量导入员工数据 → 内容查询与导出 "
                "→ 定期检查历史记录 → 需要时清除数据"
            ),
            style="CardSub.TLabel",
            justify="left",
            wraplength=1040,
        ).pack(anchor="w", pady=(6, 0))

    def _add_section(self, title, steps):
        self.sections.append(title)
        card = ttk.Frame(
            self.content, style="Card.TFrame", padding=(22, 16)
        )
        card.pack(fill="x", pady=5, padx=2)
        ttk.Label(
            card,
            text=title,
            font=(FONT_FAMILY, 12, "bold"),
            foreground="#2f6f82",
        ).pack(anchor="w")
        text = "\n".join(
            f"{index}. {step}" for index, step in enumerate(steps, 1)
        )
        ttk.Label(
            card,
            text=text,
            style="CardSub.TLabel",
            justify="left",
            wraplength=1040,
        ).pack(anchor="w", pady=(7, 0))

    def _on_mousewheel(self, event):
        if getattr(event, "num", None) == 4:
            steps = -1
        elif getattr(event, "num", None) == 5:
            steps = 1
        else:
            delta = getattr(event, "delta", 0)
            if not delta:
                return "break"
            steps = -1 if delta > 0 else 1
        self.canvas.yview_scroll(steps, "units")
        return "break"


class App(tk.Tk):
    def __init__(self, secret):
        super().__init__()
        if DateEntry is None:
            raise AppError("缺少日历组件，请重新安装程序或运行：pip install tkcalendar")
        self.title(APP_TITLE)
        self.geometry("1280x800")
        self.minsize(1000, 680)
        self.resizable(True, True)
        self.data = DataManager(os.path.join(base_dir(), DATA_FILE_NAME), secret)
        self.data.check_key()
        self.options = OptionStore(
            os.path.join(base_dir(), OPTION_STORE_FILE),
            seed_values=self.data.all_column_values(),
            legacy_dir=base_dir(),
        )
        self._focus_check_job = None
        self.bind("<FocusOut>", self._on_app_focus_out, add="+")
        self._setup_style()

        self.notebook = ttk.Frame(self)
        self.notebook.rowconfigure(0, weight=1)
        self.notebook.columnconfigure(0, weight=1)
        self.query_tab = ttk.Frame(self.notebook)
        self.add_tab = ttk.Frame(self.notebook)
        self.option_tab = OptionManagerPage(self.notebook, self)
        self.history_tab = HistoryPage(self.notebook, self)
        self.table_import_tab = TableImportPage(self.notebook, self)
        self.help_tab = HelpPage(self.notebook, self)

        container = ttk.Frame(self.query_tab)
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.pages = {}
        for key, page_cls in (
            ("query", QueryPage),
            ("result", ResultPage),
            ("team", TeamResultPage),
        ):
            page = page_cls(container, self)
            page.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = page
        self.add_tab.rowconfigure(0, weight=1)
        self.add_tab.columnconfigure(0, weight=1)
        self.pages["add"] = AddPage(self.add_tab, self)
        self.pages["add"].grid(row=0, column=0, sticky="nsew")
        self.clear_tab = ClearDataPage(self.notebook, self)
        self.tab_order = [
            self.query_tab,
            self.add_tab,
            self.option_tab,
            self.clear_tab,
            self.history_tab,
            self.table_import_tab,
            self.help_tab,
        ]
        self.tab_labels = [
            "内容查询",
            "添加生成数据",
            "新内容添加",
            "清除数据",
            "历史记录",
            "表格录入",
            "帮助",
        ]
        for tab in self.tab_order:
            tab.grid(row=0, column=0, sticky="nsew")
        self.tab_bar = PillTabBar(
            self,
            self.tab_labels,
            self._on_pill_tab,
        )
        self.tab_bar.pack(fill="x", padx=16, pady=(8, 0))
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        bind_clipboard_shortcuts(self)
        bind_combobox_behavior(self)
        self._center()
        self.show("query")

    def _setup_style(self):
        page_bg = "#eef8f4"
        card_bg = "#f9fdfc"
        text_main = "#27434b"
        text_sub = "#68818a"
        text_label = "#345861"
        primary = "#3f879d"
        primary_dark = "#2f6f82"
        border = "#c5dde1"
        blue_section = "#e8f5fb"
        green_section = "#edf8f1"

        self.configure(bg=page_bg)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=(FONT_FAMILY, 11))
        style.configure("TFrame", background=page_bg)
        style.configure("Card.TFrame", background=card_bg)
        style.layout(
            "HiddenTabs.TNotebook",
            [("Notebook.client", {"sticky": "nswe"})],
        )
        style.configure("HiddenTabs.TNotebook", background=page_bg, borderwidth=0)
        style.configure(
            "TNotebook",
            background=page_bg,
            borderwidth=0,
            tabmargins=(8, 8, 8, 0),
        )
        style.configure(
            "TNotebook.Tab",
            font=(FONT_FAMILY, 12, "bold"),
            padding=(24, 10),
            background=green_section,
            foreground=text_label,
            bordercolor=border,
        )
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", blue_section),
                ("active", "#e1f2ec"),
            ],
            foreground=[("selected", primary_dark)],
        )
        style.configure("TLabel", background=card_bg, foreground=text_main)
        style.configure(
            "Title.TLabel",
            font=(FONT_FAMILY, 20, "bold"),
            foreground=text_main,
            background=page_bg,
        )
        style.configure(
            "Sub.TLabel",
            font=(FONT_FAMILY, 10),
            foreground=text_sub,
            background=page_bg,
        )
        style.configure(
            "Field.TLabel",
            font=(FONT_FAMILY, 11, "bold"),
            foreground=text_label,
            background=blue_section,
        )
        style.configure(
            "CardSub.TLabel",
            font=(FONT_FAMILY, 10),
            foreground=text_sub,
            background=card_bg,
        )
        style.configure("BlueSection.TFrame", background=blue_section)
        style.configure(
            "BlueTitle.TLabel",
            font=(FONT_FAMILY, 12, "bold"),
            foreground=primary_dark,
            background=blue_section,
        )
        style.configure(
            "GreenSection.TFrame",
            background=green_section,
        )
        style.configure(
            "GreenTitle.TLabel",
            font=(FONT_FAMILY, 12, "bold"),
            foreground=primary_dark,
            background=green_section,
        )
        style.configure(
            "GreenField.TLabel",
            font=(FONT_FAMILY, 11, "bold"),
            foreground=text_label,
            background=green_section,
        )
        style.configure(
            "GreenSub.TLabel",
            font=(FONT_FAMILY, 10),
            foreground=text_sub,
            background=green_section,
        )
        style.configure(
            "BlueQuerySection.TFrame",
            background="#e3f3f9",
        )
        style.configure(
            "BlueQueryTitle.TLabel",
            font=(FONT_FAMILY, 12, "bold"),
            foreground=primary_dark,
            background="#e3f3f9",
        )
        style.configure(
            "BlueQueryField.TLabel",
            font=(FONT_FAMILY, 11, "bold"),
            foreground=text_label,
            background="#e3f3f9",
        )

        style.configure(
            "TButton",
            font=(FONT_FAMILY, 11),
            padding=(16, 8),
            background=card_bg,
            foreground=text_main,
            bordercolor=border,
            focuscolor=card_bg,
        )
        style.map(
            "TButton",
            background=[("pressed", "#e2e6ee"), ("active", "#f1f5fb")],
            bordercolor=[("active", "#b8c4d8")],
        )
        style.configure(
            "Primary.TButton",
            font=(FONT_FAMILY, 11, "bold"),
            padding=(24, 9),
            background=primary,
            foreground="#ffffff",
            bordercolor=primary,
            focuscolor=primary,
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", primary_dark), ("active", "#4d99ad")],
            bordercolor=[("pressed", primary_dark), ("active", "#4d99ad")],
        )
        style.configure(
            "Danger.TButton",
            font=(FONT_FAMILY, 11, "bold"),
            padding=(16, 8),
            background="#c75f5f",
            foreground="#ffffff",
            bordercolor="#b94e4e",
        )
        style.map(
            "Danger.TButton",
            background=[("pressed", "#a94646"), ("active", "#d76d6d")],
            bordercolor=[("pressed", "#a94646"), ("active", "#d76d6d")],
        )
        style.configure(
            "Option.TButton",
            font=(FONT_FAMILY, 13, "bold"),
            padding=(20, 22),
            background=blue_section,
            foreground=text_main,
            bordercolor=border,
        )
        style.map(
            "Option.TButton",
            background=[
                ("pressed", "#cfe8f1"),
                ("active", "#d9eef5"),
            ],
            bordercolor=[("active", primary)],
        )

        style.configure(
            "TCombobox",
            font=(FONT_FAMILY, 12),
            padding=6,
            arrowsize=24,
            fieldbackground=card_bg,
            background=card_bg,
            foreground=text_main,
            bordercolor=border,
            arrowcolor=text_sub,
            focuscolor=primary,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", card_bg)],
            bordercolor=[("focus", primary)],
            selectbackground=[("readonly", "#e6f4f8")],
        )
        style.configure(
            "TEntry",
            font=(FONT_FAMILY, 12),
            fieldbackground=card_bg,
            foreground=text_main,
            bordercolor=border,
            padding=5,
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", primary)],
            selectbackground=[("focus", primary)],
            selectforeground=[("focus", "#ffffff")],
        )

        style.configure(
            "TCheckbutton",
            font=(FONT_FAMILY, 11),
            background=card_bg,
            foreground=text_label,
        )
        style.map("TCheckbutton", background=[("active", card_bg)])

        style.configure(
            "Treeview",
            rowheight=32,
            font=(FONT_FAMILY, 11),
            background=card_bg,
            fieldbackground=card_bg,
            foreground=text_main,
        )
        style.configure(
            "Treeview.Heading",
            font=(FONT_FAMILY, 11, "bold"),
            background="#5a9caf",
            foreground="#ffffff",
            padding=(6, 7),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#4c8da0")],
            foreground=[("active", "#ffffff")],
        )
        style.map(
            "Treeview",
            background=[("selected", primary)],
            foreground=[("selected", "#ffffff")],
        )
        style.configure(
            "TScrollbar",
            background="#cfd8e6",
            troughcolor="#f1f4f9",
            bordercolor="#f1f4f9",
            arrowsize=12,
        )

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def report_callback_exception(self, exc, val, tb):
        _write_error_log("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror(APP_TITLE, f"程序发生错误：\n{val}")

    def _on_app_focus_out(self, event):
        if self._focus_check_job is None:
            self._focus_check_job = self.after(
                20, self._close_popdowns_if_inactive
            )

    def _close_popdowns_if_inactive(self):
        self._focus_check_job = None
        if sys.platform == "win32":
            try:
                import ctypes

                user32 = ctypes.windll.user32
                foreground = user32.GetForegroundWindow()
                process_id = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(
                    foreground, ctypes.byref(process_id)
                )
                if process_id.value != os.getpid():
                    close_combobox_popdowns(self)
                return
            except Exception:
                pass
        try:
            focused = self.focus_displayof()
        except (tk.TclError, KeyError):
            focused = None
        if focused is None or focused.winfo_toplevel() is not self:
            close_combobox_popdowns(self)

    def _on_pill_tab(self, index):
        if index == 0:
            self.show("query")
        elif index == 1:
            self.show("add")
        elif index == 2:
            self._select_tab(self.option_tab)
            self.option_tab.on_show()
        elif index == 3:
            self._select_tab(self.clear_tab)
            self.clear_tab.on_show()
        elif index == 4:
            self._select_tab(self.history_tab)
            self.history_tab.on_show()
        elif index == 5:
            self._select_tab(self.table_import_tab)
        elif index == 6:
            self._select_tab(self.help_tab)

    def _select_tab(self, tab):
        tab.tkraise()
        try:
            self.tab_bar.select(self.tab_order.index(tab))
        except (AttributeError, ValueError, tk.TclError):
            pass

    def show(self, key):
        close_combobox_popdowns(self)
        if key == "add":
            self._select_tab(self.add_tab)
        else:
            self._select_tab(self.query_tab)
        page = self.pages[key]
        page.tkraise()
        page.on_show()
        page_path = str(page) + "."
        for widget in _walk_widgets(self):
            if isinstance(widget, FilterCombobox):
                widget._popup_allowed = str(widget).startswith(page_path)
        self.after_idle(lambda: close_combobox_popdowns(self))


def run_selftest():
    tmp = tempfile.mkdtemp(prefix="emp_test_")
    try:
        path = os.path.join(tmp, DATA_FILE_NAME)
        dm = DataManager(path, "testpass")
        assert os.path.exists(path), "数据文件未创建"
        raw = open(path, "rb").read()
        assert raw.startswith(MAGIC), "数据文件未加密"
        try:
            decrypt_bytes("wrongpass", raw)
            raise AssertionError("错误密钥不应解密成功")
        except AppError:
            pass
        try:
            DataManager(path, "wrongpass").rows()
            raise AssertionError("错误密钥不应能读取数据")
        except AppError:
            pass
        assert dm.columns() == DEFAULT_COLUMNS, "默认列不正确"
        assert dm.names() == [], "初始姓名列表应为空"
        assert dm.rows() == [], "预填公式的空行不应被读取"

        dm.add_row(
            {
                "员工姓名": "张三",
                "日期": "2026-08-01",
                "工号": "E001",
                "是否新人": "否",
                "是否支援人员": "是",
                "班组名称": "甲班",
                "产品客户": "天弘",
                "所属机种": "点胶机",
                "物料代码": "BAT-1",
                "所属产线/工序": "点胶",
                "模具号": "M-100",
                "当日计划产量": "100",
                "实际产量": "95",
                "合格品数量": "90",
                "不良品数量": "5",
                "工时": "8",
                "备注": "",
            }
        )
        dm.add_row(
            {
                "员工姓名": "张三",
                "日期": "2026/8/3",
                "工号": "E001",
                "是否新人": "否",
                "班组名称": "甲班",
                "产品客户": "天弘",
                "所属机种": "点胶机",
                "物料代码": "BAT-1",
                "所属产线/工序": "点胶",
                "当日计划产量": "100",
                "实际产量": "98",
                "合格品数量": "96",
                "不良品数量": "2",
                "工时": "7",
                "备注": "加班",
            }
        )
        dm.add_row(
            {
                "员工姓名": "李四",
                "日期": "2026.08.02",
                "工号": "E002",
                "是否新人": "否",
                "班组名称": "乙班",
                "产品客户": "科勒",
                "所属机种": "齿轮箱",
                "物料代码": "GEAR-2",
                "所属产线/工序": "组装",
                "当日计划产量": "50",
                "实际产量": "45",
                "合格品数量": "44",
                "不良品数量": "1",
                "工时": "6",
                "备注": "",
            }
        )
        dm.add_row(
            {
                "员工姓名": "王五",
                "日期": "2026-08-02",
                "工号": "E003",
                "是否新人": "是",
                "班组名称": "",
                "当日计划产量": "",
                "实际产量": "",
                "合格品数量": "",
                "不良品数量": "",
                "工时": "",
                "备注": "",
            }
        )

        assert dm.names() == ["张三", "李四", "王五"], f"姓名列表错误: {dm.names()}"
        assert dm.teams() == ["乙班", "甲班"], f"班组列表错误: {dm.teams()}"
        col_vals = dm.all_column_values()
        assert col_vals["员工姓名"] == ["张三", "李四", "王五"], col_vals["员工姓名"]
        assert col_vals["班组名称"] == ["乙班", "甲班"], col_vals["班组名称"]
        rows_zs = dm.rows(name="张三")
        assert len(rows_zs) == 2, "张三应有2条记录"
        assert rows_zs[0]["日期"] == "2026-08-01", "日期未规范化"
        assert rows_zs[1]["日期"] == "2026-08-03", "斜杠日期未规范化"
        rows_team = dm.rows(team="甲班")
        assert len(rows_team) == 2, "甲班应有2条记录"
        rows_team_filtered = dm.rows(team="甲班", start="2026-08-01", end="2026-08-02")
        assert len(rows_team_filtered) == 1, "班组日期筛选错误"
        rows_filtered = dm.rows(name="张三", start="2026-08-01", end="2026-08-02")
        assert len(rows_filtered) == 1, "日期筛选错误"
        rows_all = dm.rows(start="2026-08-02", end="2026-08-03")
        assert len(rows_all) == 3, f"区间查询错误: {len(rows_all)}"

        dm.add_row(
            {
                "员工姓名": "张三",
                "日期": "2026-08-05",
                "工号": "E001",
                "是否新人": "否",
                "班组名称": "甲班",
                "产品客户": "天弘",
                "所属机种": "点胶机",
                "物料代码": "BAT-1",
                "所属产线/工序": "点胶",
                "当日计划产量": "100",
                "实际产量": "100",
                "合格品数量": "99",
                "不良品数量": "1",
                "工时": "9",
                "备注": "",
            }
        )
        assert len(dm.rows(name="张三")) == 3, "追加保存失败"
        assert len(dm.rows(team="甲班")) == 3, "班组追加失败"
        assert len(dm.rows(material="BAT-1")) == 3, "按物料代码查询失败"
        assert len(dm.rows(emp_id="E001")) == 3, "按工号查询失败"
        assert len(dm.rows(machine="点胶机")) == 3, "按所属机种查询失败"
        assert len(dm.rows(name="张三", emp_id="E001")) == 3, "姓名+工号组合查询失败"
        assert len(dm.rows(customer="科勒")) == 1, "按产品客户查询失败"
        assert len(dm.rows(line="点胶")) == 3, "按产线/工序查询失败"
        assert len(dm.rows(newbie="是")) == 1, "新人查询失败"
        assert len(dm.rows(support="是")) == 1, "支援人员查询失败"
        assert len(dm.rows(mold="M-100")) == 1, "按模具号查询失败"
        assert (
            len(
                dm.rows(
                    name="张三",
                    customer="天弘",
                    line="点胶",
                    material="BAT-1",
                )
            )
            == 3
        ), "组合查询失败"

        rate_row = dm.rows(name="张三")[0]
        assert rate_row["达成率"] == "95.0%", rate_row["达成率"]
        assert rate_row["合格率"] == "94.7%", rate_row["合格率"]

        import openpyxl

        raw_now = open(path, "rb").read()
        raw_wb = openpyxl.load_workbook(io.BytesIO(decrypt_bytes("testpass", raw_now)))
        raw_ws = raw_wb["员工数据"]
        raw_headers = [c.value for c in raw_ws[1]]
        rate_col = raw_headers.index("达成率") + 1
        quality_col = raw_headers.index("合格率") + 1
        assert str(raw_ws.cell(row=2, column=rate_col).value).startswith("=IF"), "达成率未写入公式"
        assert str(raw_ws.cell(row=2, column=quality_col).value).startswith(
            "=IF"
        ), "合格率未写入公式"
        raw_ws.cell(row=2, column=rate_col).value = "88.8%"
        buf = io.BytesIO()
        raw_wb.save(buf)
        with open(path, "wb") as f:
            f.write(encrypt_bytes("testpass", buf.getvalue()))
        assert dm.rows(name="张三")[0]["达成率"] == "88.8%", "手动填写的数值被覆盖"

        assert (
            dm.delete_rows(
                name="张三", emp_id="E001", start="2026-08-01", end="2026-08-02", preview=True
            )
            == 1
        )
        assert len(dm.rows(name="张三")) == 3, "预览不应删除数据"
        assert (
            dm.delete_rows(name="张三", emp_id="E001", start="2026-08-01", end="2026-08-02") == 1
        )
        assert len(dm.rows(name="张三")) == 2, "按日期删除失败"
        assert dm.delete_rows(team="甲班") == 2, "按班组删除数量错误"
        assert len(dm.rows(team="甲班")) == 0, "班组删除后仍有余留"
        assert dm.teams() == ["乙班"], f"班组列表未更新: {dm.teams()}"
        remaining = dm.rows(name="李四")
        assert remaining and remaining[0]["达成率"] == "90.0%", "删除后公式未修正"
        raw_wb3 = openpyxl.load_workbook(io.BytesIO(decrypt_bytes("testpass", open(path, "rb").read())))
        raw_ws3 = raw_wb3["员工数据"]
        assert str(raw_ws3.cell(row=2, column=rate_col).value).startswith(
            "=IF"
        ), "删除后公式未修复"
        assert dm.delete_rows(emp_id="E002", preview=True) == 1
        assert dm.delete_rows(emp_id="E002") == 1
        assert len(dm.rows(name="李四")) == 0, "按工号删除失败"
        rows_left = dm.rows()
        assert rows_left and "_row" in rows_left[0], "查询结果缺少行号"
        target = rows_left[0]["_row"]
        assert dm.delete_rows_by_indexes([target]) == 1, "按行号删除失败"
        assert dm.rows() == [], "按行号删除后应无数据"

        assert calc_rates("100", "95", "90", "5") == ("95.0%", "94.7%")
        assert calc_rates("0", "0", "0", "0") == ("—", "—")
        assert calc_rates("", "", "", "") == ("—", "—")
        assert calc_rates("200", "180", None, None) == ("90.0%", "—")
        assert calc_rates(None, None, "90", "10") == ("—", "90.0%")
        print("SELFTEST PASS")
        return 0
    finally:
        import shutil

        shutil.rmtree(tmp, ignore_errors=True)


def acquire_single_instance():
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(
            None, False, "Local\\NewPersonnelManagementApp"
        )
        if not handle:
            return None
        if kernel32.GetLastError() == 183:
            return None
        return handle
    except Exception:
        return object()


def _show_fatal_error(exc_type, exc_value, exc_tb):
    if exc_type is not None:
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    else:
        detail = str(exc_value)
    _write_error_log(detail)
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            APP_TITLE,
            f"程序发生错误：\n{exc_value}\n\n详细信息已记录到“程序错误日志.txt”",
        )
        root.destroy()
    except Exception:
        pass


def _write_error_log(text):
    try:
        log_path = os.path.join(base_dir(), "程序错误日志.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}]\n{text}\n")
    except Exception:
        pass


def state_path():
    return os.path.join(base_dir(), STATE_FILE)


def load_state():
    try:
        with open(state_path(), "rb") as f:
            raw = f.read()
        data = decrypt_bytes(RENEWAL_KEY, raw)
        return json.loads(data.decode("utf-8"))
    except Exception:
        return {}


def save_state(state):
    payload = json.dumps(state, ensure_ascii=False).encode("utf-8")
    enc = encrypt_bytes(RENEWAL_KEY, payload)
    tmp = state_path() + ".tmp"
    with open(tmp, "wb") as f:
        f.write(enc)
    os.replace(tmp, state_path())


def show_message(title, text):
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title, text)
        root.destroy()
    except Exception:
        pass


def ask_renewal_key():
    result = {}
    root = tk.Tk()
    root.title("解锁程序")
    root.resizable(False, False)
    root.configure(bg="#ffffff")
    frame = ttk.Frame(root, padding=24)
    frame.pack()
    ttk.Label(
        frame,
        text=(
            f"使用期限已结束（{LICENSE_EXPIRE_DATE:%Y-%m-%d}），"
            "请输入到期解锁密钥："
        ),
        font=(FONT_FAMILY, 11),
    ).pack(anchor="w")
    var = tk.StringVar()
    entry = ttk.Entry(frame, textvariable=var, show="*", width=26)
    entry.pack(pady=(10, 0))
    entry.focus_set()

    def submit():
        result["value"] = var.get()
        root.destroy()

    def cancel():
        result["value"] = None
        root.destroy()

    btns = ttk.Frame(frame)
    btns.pack(pady=(12, 0))
    ttk.Button(btns, text="确 定", command=submit).pack(side="left", padx=(0, 10))
    ttk.Button(btns, text="退 出", command=cancel).pack(side="left")
    entry.bind("<Return>", lambda e: submit())
    root.bind("<Escape>", lambda e: cancel())

    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 3
    root.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    root.mainloop()
    return result.get("value")


def check_license():
    """返回 None 表示可正常使用；否则返回需要提示并退出的文本。"""
    state = load_state()

    if state.get("locked_until"):
        locked = datetime.datetime.fromisoformat(state["locked_until"])
        now = datetime.datetime.now()
        if locked > now:
            remain = locked - now
            hours = int(remain.total_seconds() // 3600)
            mins = int((remain.total_seconds() % 3600) // 60)
            return f"密钥错误次数过多，程序已锁定。\n请等待 {hours} 小时 {mins} 分钟后重试。"
        state["attempts"] = 0
        state["locked_until"] = None
        save_state(state)

    first_run = state.get("first_run")
    if not first_run:
        state["first_run"] = datetime.date.today().isoformat()
        state["attempts"] = 0
        state["locked_until"] = None
        state["unlocked"] = False
        save_state(state)
        return None

    if state.get("unlocked"):
        return None

    expire_date = LICENSE_EXPIRE_DATE
    if datetime.date.today() <= expire_date:
        return None

    attempts = int(state.get("attempts", 0))
    while attempts < MAX_ATTEMPTS:
        renewal_key = ask_renewal_key()
        if renewal_key is None:
            return "已取消解锁，程序退出。"
        if renewal_key.strip() == RENEWAL_KEY:
            state["unlocked"] = True
            save_state(state)
            return None
        attempts += 1
        state["attempts"] = attempts
        save_state(state)
        if attempts >= MAX_ATTEMPTS:
            state["locked_until"] = (
                datetime.datetime.now() + datetime.timedelta(hours=LOCK_HOURS)
            ).isoformat()
            state["attempts"] = 0
            save_state(state)
            return f"密钥错误已达 {MAX_ATTEMPTS} 次，程序已锁定。\n请 24 小时后再试。"
        messagebox.showwarning(
            "密钥错误", f"密钥错误，还剩 {MAX_ATTEMPTS - attempts} 次机会。"
        )
    return "解锁失败，程序退出。"


def main():
    mutex = acquire_single_instance()
    if mutex is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(APP_TITLE, "程序已经在运行中，请查看屏幕上的程序窗口。")
        root.destroy()
        return
    sys.excepthook = _show_fatal_error
    try:
        license_error = check_license()
    except Exception as exc:
        license_error = f"授权检查失败：{exc}"
    if license_error:
        show_message("程序无法使用", license_error)
        return
    try:
        app = App(PROGRAM_KEY)
    except AppError as exc:
        _show_fatal_error(None, exc, None)
        return
    app.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(run_selftest())
    if "--smoke" in sys.argv:
        enable_dpi_awareness()
        app = App(PROGRAM_KEY)
        if len(app.tab_order) != 7:
            raise AppError("选项卡数量不正确")
        tab_names = list(app.tab_labels)
        if tab_names != [
            "内容查询",
            "添加生成数据",
            "新内容添加",
            "清除数据",
            "历史记录",
            "表格录入",
            "帮助",
        ]:
            raise AppError("选项卡名称或顺序不正确")
        if not hasattr(app, "help_tab"):
            raise AppError("帮助选项卡未正确加载")
        app._select_tab(app.help_tab)
        app.update_idletasks()
        if len(app.help_tab.sections) != 7:
            raise AppError("帮助选项卡操作流程不完整")
        if not any(
            "日常录入" in title for title in app.help_tab.sections
        ):
            raise AppError("帮助选项卡缺少日常录入流程")
        if any(
            "密钥" in title for title in app.help_tab.sections
        ):
            raise AppError("帮助选项卡不应显示密钥解锁流程")
        if len(app.option_tab.buttons) != len(OPTION_FIELDS):
            raise AppError("内容添加按钮数量不正确")
        app._select_tab(app.table_import_tab)
        app.update_idletasks()
        import_buttons = [
            str(widget.cget("text"))
            for widget in _walk_widgets(app.table_import_tab)
            if isinstance(widget, ttk.Button)
        ]
        if import_buttons != ["生成Excel模板", "导入Excel表格"]:
            raise AppError("表格录入页按钮布局不正确")
        app.history_tab.on_show()
        app.show("add")
        app.update_idletasks()
        support_var = app.pages["add"].field_widgets.get("是否支援人员")
        if support_var is None or support_var.get() != "否":
            raise AppError("“是否支援人员”字段未正确加载")
        if "模具号" not in app.pages["add"].field_combos:
            raise AppError("“模具号”添加字段未正确加载")
        if not hasattr(app.pages["query"], "mold_combo"):
            raise AppError("“模具号”查询字段未正确加载")
        if not hasattr(app.pages["query"], "query_support"):
            raise AppError("支援人员查询功能未正确加载")
        if not hasattr(app.pages["query"], "query_date"):
            raise AppError("指定日期查询功能未正确加载")
        if app.pages["query"].single_date_entry.get_date() is None:
            raise AppError("指定日期选择框未正确加载")
        filter_combo = app.pages["add"].name_combo
        filter_combo.set_options(("王小明", "王丽", "李四"))
        filter_combo.set("王")
        filter_combo._typing = True
        filter_combo._apply_filter()
        app.update_idletasks()
        if tuple(filter_combo["values"]) != ("王小明", "王丽"):
            raise AppError("下拉框输入筛选未能正确匹配")
        filter_combo.tk.call("ttk::combobox::Unpost", filter_combo._w)
        filter_combo.clear_selection()
        date_entry = app.pages["add"].date_entry
        date_entry.drop_down()
        app.update_idletasks()
        if date_entry._top_cal is None or not date_entry._top_cal.winfo_exists():
            raise AppError("日历下拉窗口未能打开")
        app.after(500, app.destroy)
        app.mainloop()
        print("SMOKE OK")
        sys.exit(0)
    enable_dpi_awareness()
    main()
