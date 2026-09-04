"""Tkinter 入口模块，负责界面、后台线程与批量处理编排。"""

import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

from config import MODEL_MAP
from excel_export import (
    export_excel,
    get_last_export_dir,
    get_output_dir,
    save_last_export_dir,
)
from feishu_client import get_tenant_access_token, send_to_bitable
from logging_utils import log_queue, print_log
from mock_data import generate_mock_data
from ocr_client import (
    OCRAborted,
    OCRResultTimeout,
    call_get_result_api,
    call_process_api,
    upload_file_to_server,
)
from parsers import (
    get_core_headers,
    get_default_order_type_label,
    get_order_type_labels,
    get_preview_layout,
    merge_preview_rows,
    parse_commit_result,
)
from wms_client import (
    build_put_original_sales_order_payload,
    build_put_purchase_order_payload,
    format_wms_response,
    is_wms_send_success,
    send_put_original_sales_order,
    send_put_purchase_order,
)


ui_message_queue = queue.Queue()
worker_thread = None
export_thread = None
continue_thread = None
abort_event = threading.Event()
preview_select_text = ""
preview_files = []
active_tree = None
continue_query_active = False
last_combo_text = ""
progress_percent = 0
progress_color = "#16A34A"
session_log_lines = []
log_window = None
log_window_text = None
wms_thread = None
wms_send_active = False
wms_window = None
wms_window_request_text = None
wms_window_response_text = None
wms_confirm_button = None

if sys.platform == "win32":
    PREVIEW_HEADING_FONT = ("Microsoft YaHei UI", 10, "bold")
else:
    PREVIEW_HEADING_FONT = ("黑体", 15, "bold")
MAX_BATCH_FILES = 5
MANUAL_STATUS = "人工填写"
MANUAL_FILENAME = "空白单据"
OSCAR_HEADER_DISPLAY_LABELS = {
    "供应商": "客户/供应商",
}


class EditableTreeview(ttk.Treeview):
    """支持双击编辑单元格、插入/新增/删除行以及整行复制粘贴的 Treeview。"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._entry = None
        self._bound_item = None
        self._bound_col = None
        self._clipboard = []
        self.bind("<Double-1>", self._start_edit)
        for sequence in ("<Control-c>", "<Command-c>"):
            self.bind(sequence, self._copy_shortcut)
        for sequence in ("<Control-v>", "<Command-v>"):
            self.bind(sequence, self._paste_shortcut)

    def _start_edit(self, event):
        if self.identify("region", event.x, event.y) != "cell":
            return
        row_id = self.identify_row(event.y)
        col_index = int(self.identify_column(event.x).replace("#", "")) - 1
        headers = self["columns"]
        if not row_id or col_index < 0 or col_index >= len(headers):
            return
        if self._is_summary_row(row_id):
            return
        bbox = self.bbox(row_id, self.identify_column(event.x))
        if not bbox:
            return
        self._finish_edit()
        col_id = headers[col_index]
        entry = tk.Entry(self)
        entry.insert(0, self.set(row_id, col_id))
        entry.select_range(0, tk.END)
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.focus_set()
        self._entry = entry
        self._bound_item = row_id
        self._bound_col = col_id
        entry.bind("<Return>", self._finish_edit)
        entry.bind("<FocusOut>", self._finish_edit)
        entry.bind("<Escape>", self._cancel_edit)

    def _finish_edit(self, _event=None):
        if self._entry is not None:
            self.set(self._bound_item, self._bound_col, self._entry.get())
            self._destroy_edit()

    def _cancel_edit(self, _event=None):
        self._destroy_edit()

    def _destroy_edit(self):
        if self._entry is not None:
            self._entry.destroy()
            self._entry = None
            self._bound_item = None
            self._bound_col = None

    def _copy_shortcut(self, _event=None):
        active_tree_copy_selected()
        return "break"

    def _paste_shortcut(self, _event=None):
        active_tree_paste_row()
        return "break"

    def _is_summary_row(self, row_id):
        return "summary_row" in self.item(row_id, "tags")

    def _exportable_rows_in_order(self):
        ordered = []

        def visit(parent_id):
            for row_id in self.get_children(parent_id):
                if not parent_id and self._is_summary_row(row_id):
                    visit(row_id)
                else:
                    ordered.append(row_id)

        visit("")
        return ordered

    def _selected_rows_in_order(self):
        selected = set(self.selection())
        if not selected:
            return []
        return [
            row_id for row_id in self._exportable_rows_in_order()
            if row_id in selected
        ]

    def has_clipboard(self):
        return bool(self._clipboard)

    def has_copyable_selection(self):
        return bool(self._selected_rows_in_order())

    def copy_selected(self):
        selected = self._selected_rows_in_order()
        if not selected:
            self._clipboard = []
            return False
        headers = self["columns"]
        self._clipboard = [
            [self.set(row_id, col) for col in headers]
            for row_id in selected
        ]
        return True

    def clear_clipboard(self):
        self._clipboard.clear()

    def insert_row_after_selection(self):
        if not self["columns"]:
            return None
        selected = self._selected_rows_in_order()
        if selected:
            target = selected[-1]
        elif self.selection():
            target = self.selection()[-1]
        else:
            target = None
        if target is not None:
            parent_id = self.parent(target)
            children = self.get_children(parent_id)
            index = children.index(target) + 1
        else:
            parent_id = ""
            index = tk.END
        row_id = self.insert(
            parent_id, index, values=("",) * len(self["columns"]),
            tags=("new_row",)
        )
        if parent_id:
            self.item(parent_id, open=True)
        self._renumber()
        self.selection_set(row_id)
        self.see(row_id)
        return row_id

    def paste_clipboard(self):
        if not self._clipboard or not self["columns"]:
            return False
        selected = self._selected_rows_in_order()
        if selected:
            target = selected[-1]
        elif self.selection():
            target = self.selection()[-1]
        else:
            target = None
        if target is not None:
            parent_id = self.parent(target)
            children = self.get_children(parent_id)
            index = children.index(target) + 1
        else:
            parent_id = ""
            index = len(self.get_children())
        pasted = []
        for offset, values in enumerate(self._clipboard):
            row_id = self.insert(
                parent_id, index + offset, values=list(values),
                tags=("new_row",)
            )
            pasted.append(row_id)
        if parent_id:
            self.item(parent_id, open=True)
        self._renumber()
        self.selection_set(*pasted)
        self.see(pasted[-1])
        return True

    def add_row(self):
        if self["columns"]:
            self.insert("", tk.END, values=("",) * len(self["columns"]),
                        tags=("new_row",))
            self._renumber()
            self.see(self.get_children()[-1])

    def delete_selected(self):
        for row_id in self.selection():
            try:
                self.delete(row_id)
            except tk.TclError:
                pass
        self._renumber()

    def _renumber(self):
        for index, row_id in enumerate(self.get_children(), 1):
            self.item(row_id, text=index)
            if self._is_summary_row(row_id):
                for child_index, child_id in enumerate(
                    self.get_children(row_id), 1
                ):
                    self.item(child_id, text=f"{index}.{child_index}")

    def auto_size_columns(self, max_width=320, min_width=90, padding=26,
                          fill_width=False):
        """按内容自动收缩列宽并居中，长内容由用户拖动列宽查看。"""
        font = tkfont.nametofont("TkDefaultFont")
        heading_font = tkfont.Font(font=PREVIEW_HEADING_FONT)
        for col in self["columns"]:
            content_w = []

            def collect_values(parent_id):
                for row_id in self.get_children(parent_id):
                    content_w.append(font.measure(self.set(row_id, col)))
                    collect_values(row_id)

            collect_values("")
            width = max([heading_font.measure(col)] + content_w) + padding
            width = min(max(min_width, width), max_width)
            self.column(col, width=width, minwidth=min_width,
                        stretch=fill_width, anchor="center")
            self.heading(col, anchor="center")

    def clear_rows(self):
        for row_id in self.get_children():
            self.delete(row_id)

    def get_data(self):
        headers = list(self["columns"])
        rows = []

        def visit(parent_id):
            for row_id in self.get_children(parent_id):
                if not parent_id and self._is_summary_row(row_id):
                    visit(row_id)
                else:
                    rows.append([
                        self.set(row_id, col) for col in headers
                    ])

        visit("")
        return headers, rows


def _append_line_to_log_window(line):
    """把一行日志追加到已打开的日志窗口，避免跨线程操作 Tkinter。"""
    if log_window is None or log_window_text is None:
        return
    try:
        if not log_window.winfo_exists():
            return
        log_window_text.config(state=tk.NORMAL)
        log_window_text.insert(tk.END, line + "\n")
        log_window_text.see(tk.END)
        log_window_text.config(state=tk.DISABLED)
    except tk.TclError:
        pass


def _refresh_log_window():
    """刷新日志窗口内容，用于打开时载入本次会话已有日志。"""
    if log_window is None or log_window_text is None:
        return
    try:
        if not log_window.winfo_exists():
            return
        log_window_text.config(state=tk.NORMAL)
        log_window_text.delete("1.0", tk.END)
        for line in session_log_lines:
            log_window_text.insert(tk.END, line + "\n")
        log_window_text.see(tk.END)
        log_window_text.config(state=tk.DISABLED)
    except tk.TclError:
        pass


def flush_log():
    """由主线程把日志队列写入会话缓冲与打开的日志窗口。"""
    while True:
        try:
            line = log_queue.get_nowait()
        except queue.Empty:
            break
        session_log_lines.append(line)
        _append_line_to_log_window(line)
    if log_window is not None and log_window_text is not None:
        try:
            log_window_text.update_idletasks()
        except tk.TclError:
            pass


def open_log_window():
    """打开或聚焦本次会话的日志窗口。"""
    global log_window, log_window_text
    flush_log()
    if log_window is not None:
        try:
            if log_window.winfo_exists():
                log_window.deiconify()
                log_window.lift()
                log_window.focus_force()
                return
        except tk.TclError:
            log_window = None
            log_window_text = None

    log_window = tk.Toplevel(win)
    log_window.title("处理日志")
    log_window.geometry("900x500")
    log_window.minsize(400, 200)

    body = tk.Frame(log_window)
    body.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    log_window_text = tk.Text(body, state=tk.DISABLED, wrap=tk.WORD)
    log_scrollbar = ttk.Scrollbar(
        body, orient="vertical", command=log_window_text.yview
    )
    log_window_text.configure(yscrollcommand=log_scrollbar.set)
    log_window_text.grid(row=0, column=0, sticky="nsew")
    log_scrollbar.grid(row=0, column=1, sticky="ns")

    body.rowconfigure(0, weight=1)
    body.columnconfigure(0, weight=1)
    log_window.protocol("WM_DELETE_WINDOW", close_log_window)
    _refresh_log_window()


def close_log_window():
    """关闭日志窗口并允许重新打开。"""
    global log_window, log_window_text
    if log_window is not None:
        log_window.destroy()
    log_window = None
    log_window_text = None


def _wms_text_pane(parent, title):
    """创建接口发送窗口中的只读文本分栏。"""
    tk.Label(parent, text=title, font=BUTTON_FONT, anchor="w").grid(
        row=0, column=0, sticky="w", pady=(0, 4)
    )
    text = tk.Text(parent, state=tk.DISABLED, wrap=tk.NONE)
    vsb = ttk.Scrollbar(parent, orient="vertical", command=text.yview)
    hsb = ttk.Scrollbar(parent, orient="horizontal", command=text.xview)
    text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    text.grid(row=1, column=0, sticky="nsew")
    vsb.grid(row=1, column=1, sticky="ns")
    hsb.grid(row=2, column=0, sticky="ew")
    parent.rowconfigure(1, weight=1)
    parent.columnconfigure(0, weight=1)
    return text


def _replace_wms_response(token, text):
    """用最新接口回告覆盖本次发送对应的二级窗口返回区。"""
    if wms_window_response_text is None or id(wms_window_response_text) != token:
        return
    try:
        if not wms_window_response_text.winfo_exists():
            return
        wms_window_response_text.config(state=tk.NORMAL)
        wms_window_response_text.delete("1.0", tk.END)
        wms_window_response_text.insert(tk.END, text)
        wms_window_response_text.yview_moveto(0)
        wms_window_response_text.config(state=tk.DISABLED)
    except tk.TclError:
        pass


def poll_log_queue():
    """周期刷新日志队列，保证日志窗口随时显示最新内容。"""
    flush_log()
    win.after(200, poll_log_queue)


# ---------------- 主流程函数 ----------------
def run_task():
    """根据当前模板启动真实或模拟批量处理。"""
    if continue_query_active:
        messagebox.showwarning(
            "温馨提示",
            "正在继续查询原任务，请等待完成后再次选择文件"
        )
        return
    select_text = combo_model.get().strip()
    if not select_text or select_text not in MODEL_MAP:
        messagebox.showwarning("温馨提示", "请先选择单据规则！")
        return

    mock_mode = mock_var.get()
    files = ()
    if not mock_mode:
        files = filedialog.askopenfilenames(
            title="选择要处理的文件",
            filetypes=[
                ("全部支持文件", "*.png;*.jpg;*.jpeg;*.pdf"),
                ("图片文件", "*.png;*.jpg;*.jpeg"),
                ("PDF 文件", "*.pdf"),
                ("PNG 文件", "*.png"),
                ("JPG 文件", "*.jpg;*.jpeg"),
                ("所有文件", "*.*")
            ]
        )
        if not files:
            return
        if len(files) > MAX_BATCH_FILES:
            messagebox.showwarning(
                "温馨提示",
                f"每次最多选择 {MAX_BATCH_FILES} 个文件，请重新选择。"
            )
            return

    clear_preview()
    abort_event.clear()
    abort_btn.config(state=tk.NORMAL)
    btn.config(text="正在处理...", state=tk.DISABLED)
    mock_check.config(state=tk.DISABLED)
    set_progress_state(50, "处理进度：已开始（50%）")
    win.update()
    current_model_id = MODEL_MAP[select_text]

    global worker_thread
    worker_thread = threading.Thread(
        target=process_batch_worker,
        args=(files, select_text, current_model_id, mock_mode),
        daemon=True
    )
    worker_thread.start()
    win.after(100, poll_ui_queue)


def abort_processing():
    """请求中止当前 OCR 处理批次或继续查询。"""
    abort_event.set()
    abort_btn.config(state=tk.DISABLED)
    if continue_query_active:
        print_log("收到中止请求，正在停止继续查询")
    else:
        print_log("收到中止请求，正在停止当前批次")
    set_progress_state(50, "处理进度：正在中止...", "#D97706")


def finish_abort_state():
    """恢复中止后的初始界面状态。"""
    abort_btn.config(state=tk.DISABLED)
    btn.config(text="选择文件并开始处理", state=tk.NORMAL)
    mock_check.config(state=tk.NORMAL)
    set_progress_state(0, "处理进度：已中止", "#D97706")


def process_batch_worker(files, select_text, current_model_id, mock_mode):
    """后台线程入口，统一捕获批量处理异常。"""
    try:
        if mock_mode:
            if abort_event.is_set():
                ui_message_queue.put(("processing_aborted",))
                return
            process_mock_batch(select_text)
        elif not process_batch(files, select_text, current_model_id):
            print_log("批次已由用户中止，正在恢复界面")
            ui_message_queue.put(("processing_aborted",))
            return
        if abort_event.is_set():
            print_log("批次已由用户中止，正在恢复界面")
            ui_message_queue.put(("processing_aborted",))
    except Exception as e:
        print_log(f"处理流程异常: {str(e)}")
        ui_message_queue.put(("processing_error", f"处理流程异常: {str(e)}"))


def _send_feishu_statistics(select_text, total_files):
    """后台发送本批次飞书统计，失败只记录日志。"""
    print_log("开始同步统计数据至飞书多维表格...")
    token = get_tenant_access_token()
    send_to_bitable(token, select_text, total_files)


def process_batch(files, select_text, current_model_id):
    """逐文件执行 OCR 处理，按文件组织预览结果并通过队列回传 UI。"""
    total_files = len(files)
    file_results = []
    core_data = []
    success_count = 0
    fail_count = 0
    pending_count = 0
    print_log(f"===== 批量处理，共{total_files}个文件，模版：{select_text} =====")

    for done, file_path in enumerate(files, 1):
        if abort_event.is_set():
            print_log("已收到中止请求，停止处理新文件")
            return False
        filename = os.path.basename(file_path)
        status = "失败"
        res_msg = ""
        file_id = ""
        file_url = ""
        req_uuid = ""
        parsed_rows = []
        split_groups = []
        ocr_result_dict = None

        try:
            file_id, file_url = upload_file_to_server(file_path)
            if abort_event.is_set():
                raise OCRAborted()
            # 传入当前选中的单据规则名称，用于内部判断sysCode
            req_uuid, _ = call_process_api(file_url, filename, file_id, current_model_id, select_text)
            if abort_event.is_set():
                raise OCRAborted()
            _, ocr_result_dict = call_get_result_api(req_uuid, abort_event)
            if abort_event.is_set():
                raise OCRAborted()

            if ocr_result_dict and ocr_result_dict.get("status") is True:
                commit_result = ocr_result_dict.get("data", {}).get("commitResult", {})
                parsed_rows, split_groups = parse_commit_result(
                    select_text, commit_result, filename
                )
                core_data.extend(parsed_rows)
                if select_text == "GE-发票单":
                    print_log(f"{filename} 处理完成后总明细行数:{len(core_data)}")

                status = "成功"
                res_msg = "处理成功"
                success_count += 1
                print_log(f"✅ [{filename}] 处理成功")

            else:
                raise Exception("OCR返回识别状态异常")

        except OCRResultTimeout:
            status = "结果未生成"
            res_msg = "识别5分钟未生成结果，可稍后点击“继续查询原任务”"
            pending_count += 1
            print_log(f"⏳ [{filename}] 5分钟内未生成识别结果，等待后续继续查询")
        except OCRAborted:
            print_log("批次已由用户中止")
            return False
        except Exception as e:
            status = "失败"
            res_msg = str(e)
            fail_count += 1
            print_log(f"❌ [{filename}] 处理失败：{e}")

        file_results.append({
            "filename": filename,
            "status": status,
            "message": res_msg,
            "rows": parsed_rows,
            "split_groups": split_groups,
            "req_uuid": req_uuid,
            "log_row": [filename, file_id, file_url, req_uuid, "", status, res_msg, ""],
        })
        ui_message_queue.put(("progress", done, total_files))

    print_log(f"\n===== 批量处理结束 =====")
    print_log(f"总文件：{total_files} | 成功：{success_count} | 失败：{fail_count} | "
              f"结果未生成：{pending_count} | 有效明细行数：{len(core_data)}")
    if abort_event.is_set():
        print_log("批次已由用户中止，不进入预览")
        return False
    threading.Thread(
        target=_send_feishu_statistics,
        args=(select_text, total_files),
        daemon=True,
    ).start()

    ui_message_queue.put(("preview", select_text,
                          get_core_headers(select_text), file_results,
                          success_count, fail_count))
    return True


def process_mock_batch(select_text):
    """生成两个演示文件页签并直接进入预览流程。"""
    print_log(f"===== 模拟数据模式，模板：{select_text} =====")
    ui_message_queue.put(("progress", 1, 2))
    headers, file_results = generate_mock_data(select_text)
    for file_result in file_results:
        file_result["log_row"] = [
            file_result["filename"], "", "", "", "",
            file_result["status"], file_result["message"], ""
        ]
    success_count = sum(item["status"] == "成功" for item in file_results)
    fail_count = len(file_results) - success_count
    total_rows = sum(len(item["rows"]) for item in file_results)
    print_log(f"生成模拟页签数: {len(file_results)}，明细总行数: {total_rows}")
    ui_message_queue.put(("progress", 2, 2))
    ui_message_queue.put(("preview", select_text, headers,
                          file_results, success_count, fail_count))


# ---------------- 预览与导出交互 ----------------
def _short_tab_label(filename, max_len=22):
    """生成长文件名页签使用的短标签。"""
    name = os.path.basename(filename)
    if len(name) <= max_len:
        return name
    return name[:max_len - 1] + "…"


def _header_values_from_rows(rows, full_headers, header_fields):
    """从头组单据行中提取每个 Header 字段的首个非空值。"""
    values = {field: "" for field in header_fields}
    for row in rows or []:
        row_map = dict(zip(full_headers, row))
        for field in header_fields:
            if not values[field] and str(row_map.get(field, "")).strip():
                values[field] = row_map[field]
    return values


def _detail_rows_from_full_rows(rows, split_groups, full_headers, detail_fields):
    """从完整数据行中提取预览明细字段，并生成拆分汇总分组元数据。"""
    detail_rows = []
    for row in rows or []:
        row_map = dict(zip(full_headers, row))
        detail_rows.append([row_map.get(field, "") for field in detail_fields])

    preview_groups = []
    for group in split_groups or []:
        child_indexes = [
            index for index in group.get("child_indexes", [])
            if index < len(detail_rows)
        ]
        if not child_indexes:
            continue
        summary_row = dict(zip(full_headers, group.get("summary_row") or []))
        summary_values = [
            summary_row.get(field, "") for field in detail_fields
        ]
        kept_fields = {"ITEM NUMBER", "QTY", "LPN Number"}
        for field_index, field in enumerate(detail_fields):
            if field not in kept_fields:
                summary_values[field_index] = ""
        preview_groups.append({
            "summary": summary_values,
            "children": [detail_rows[index] for index in child_indexes],
            "child_indexes": child_indexes,
        })
    return detail_rows, preview_groups


def _build_scrolled_preview_tree(parent, columns, rows, preview_groups=None):
    """创建带滚动条的明细预览表格，并返回其外层容器与 Treeview。"""
    frame = tk.Frame(parent)
    tree = EditableTreeview(
        frame,
        style="Preview.Treeview",
        selectmode="extended",
    )
    tree.bind("<<TreeviewSelect>>",
              lambda _event: refresh_row_action_state())
    tree.tag_configure("new_row", background="#FFF3CD")
    tree.tag_configure("zebra_even", background="#FFFFFF")
    tree.tag_configure("zebra_odd", background="#EFF5F9")
    default_font = tkfont.nametofont("TkDefaultFont")
    summary_font = (
        default_font.actual()["family"],
        default_font.actual()["size"],
        "bold",
    )
    tree.tag_configure(
        "summary_row",
        background="#FEF3C7",
        foreground="#334155",
        font=summary_font,
    )

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    tree["columns"] = columns
    tree.column("#0", width=60, minwidth=52, stretch=False, anchor="center")
    for column in columns:
        tree.heading(column, text=column)

    child_to_group = {}
    for group in preview_groups or []:
        for index in group["child_indexes"]:
            child_to_group[index] = group
    consumed = set()
    export_index = 0
    for row_index, row in enumerate(rows):
        if row_index in consumed:
            continue
        group = child_to_group.get(row_index)
        if group:
            summary_id = tree.insert(
                "", tk.END, values=group["summary"],
                tags=("summary_row",)
            )
            tree.item(summary_id, open=True)
            for child_index in group["child_indexes"]:
                tag = (
                    "zebra_even" if export_index % 2 == 0 else "zebra_odd"
                )
                tree.insert(
                    summary_id, tk.END, values=rows[child_index],
                    tags=(tag,)
                )
                export_index += 1
            consumed.update(group["child_indexes"])
            continue

        tag = "zebra_even" if export_index % 2 == 0 else "zebra_odd"
        insert_args = {"values": row}
        insert_args["tags"] = (tag,)
        tree.insert("", tk.END, **insert_args)
        export_index += 1
    tree._renumber()
    tree.auto_size_columns(fill_width=True)
    return frame, tree


def _update_preview_header(file_result, field, value):
    """保存单据头编辑结果，导出重组时供所有明细行使用。"""
    if field in file_result.get("header_values", {}):
        file_result["header_values"][field] = value


def _build_header_form(parent, file_result, header_fields, header_values, select_text):
    """把单据头字段渲染为多列表单，编辑时直接同步到导出数据。"""
    form = tk.Frame(parent)
    columns = 5 if len(header_fields) >= 5 else max(1, len(header_fields))
    for index, field in enumerate(header_fields):
        cell = tk.Frame(form)
        cell.grid(row=index // columns, column=index % columns,
                  sticky="nsew", padx=4, pady=2)
        label_fg = "#B42318" if field in ("订单类型", "运单号") else "#111827"
        display_label = OSCAR_HEADER_DISPLAY_LABELS.get(field, field)
        tk.Label(cell, text=display_label, anchor="w",
                 font=("黑体", 11, "bold"), fg=label_fg).pack(fill=tk.X)
        value_var = tk.StringVar(master=form, value=header_values.get(field, ""))
        if field == "订单类型":
            order_type_labels = get_order_type_labels(select_text)
            default_label = get_default_order_type_label(select_text)
            if default_label and value_var.get() not in order_type_labels:
                value_var.set(default_label)
            ttk.Combobox(
                cell, textvariable=value_var, state="readonly",
                values=order_type_labels,
            ).pack(fill=tk.X)
        else:
            tk.Entry(cell, textvariable=value_var).pack(fill=tk.X)
        value_var.trace_add(
            "write",
            lambda *_args, field=field, value_var=value_var,
            file_result=file_result: _update_preview_header(
                file_result, field, value_var.get()
            ),
        )
    for col_index in range(columns):
        form.columnconfigure(col_index, weight=1, uniform="header")
    return form


def _build_file_tab(file_result, headers, header_fields, detail_fields,
                    select_text, add_tab=True):
    """为单个文件创建“单据头表单 + 明细”预览页签，并返回明细表格。"""
    tab = ttk.Frame(preview_notebook)
    tab.pack_propagate(False)
    if add_tab:
        preview_notebook.add(tab, text=_short_tab_label(file_result["filename"]))

    status_text = f"状态：{file_result['status']}"
    if not file_result.get("rows"):
        status_text += "；当前无明细数据"
    if file_result.get("message"):
        status_text += f"；{file_result['message']}"
    status_label = tk.Label(
        tab, text=status_text, anchor="w",
        fg=("#D97706" if file_result["status"] == "结果未生成"
            else "#B42318" if file_result["status"] == "失败" else "#111827")
    )
    status_label.pack(fill=tk.X, padx=8, pady=(0, 4))

    file_result["status_label"] = status_label
    header_values = _header_values_from_rows(
        file_result.get("rows") or [], headers, header_fields
    )
    if not header_values.get("订单类型"):
        prior_order_type = file_result.get("header_values", {}).get("订单类型", "")
        if prior_order_type in get_order_type_labels(select_text):
            header_values["订单类型"] = prior_order_type
        else:
            default_label = get_default_order_type_label(select_text)
            if default_label:
                header_values["订单类型"] = default_label
    file_result["header_values"] = header_values

    tk.Label(tab, text="单据头", anchor="w",
             font=("黑体", 12, "bold")).pack(fill=tk.X, padx=8, pady=(0, 2))
    header_panel = _build_header_form(
        tab, file_result, header_fields, header_values, select_text
    )
    header_panel.pack(fill=tk.X, padx=8, pady=(0, 4))

    tk.Label(tab, text="明细", anchor="w",
             font=("黑体", 12, "bold")).pack(fill=tk.X, padx=8, pady=(0, 2))
    detail_rows, preview_groups = _detail_rows_from_full_rows(
        file_result.get("rows") or [], file_result.get("split_groups") or [],
        headers, detail_fields
    )
    detail_panel, detail_tree = _build_scrolled_preview_tree(
        tab, detail_fields, detail_rows, preview_groups
    )
    detail_panel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

    file_result["tree"] = detail_tree
    file_result["tab"] = tab
    return detail_tree


def _refresh_file_status_label(file_result):
    """按文件最新状态刷新页签顶部状态文案。"""
    status_label = file_result.get("status_label")
    if status_label is None:
        return
    status_text = f"状态：{file_result['status']}"
    if not file_result.get("rows"):
        status_text += "；当前无明细数据"
    if file_result.get("message"):
        status_text += f"；{file_result['message']}"
    status_label.config(
        text=status_text,
        fg=("#D97706" if file_result["status"] == "结果未生成"
            else "#B42318" if file_result["status"] == "失败" else "#111827"),
    )


def _replace_file_tab(file_result, headers, header_fields, detail_fields,
                      select_text):
    """重建单个文件页签，保留原页签位置与选中状态。"""
    global active_tree
    old_tab = file_result["tab"]
    was_selected = str(preview_notebook.select()) == str(old_tab)
    index = preview_notebook.index(old_tab)
    active_tree = None
    old_tab.destroy()
    _build_file_tab(
        file_result, headers, header_fields, detail_fields, select_text,
        add_tab=False,
    )
    if preview_notebook.tabs():
        preview_notebook.insert(
            index,
            file_result["tab"],
            text=_short_tab_label(file_result["filename"]),
        )
        if was_selected:
            preview_notebook.select(file_result["tab"])
    else:
        preview_notebook.add(
            file_result["tab"],
            text=_short_tab_label(file_result["filename"]),
        )
        preview_notebook.select(file_result["tab"])
    on_preview_tab_changed()


def _background_task_active():
    """返回是否有后台处理、导出或续查线程正在运行。"""
    return (
        (worker_thread is not None and worker_thread.is_alive())
        or (export_thread is not None and export_thread.is_alive())
        or (continue_thread is not None and continue_thread.is_alive())
        or (wms_thread is not None and wms_thread.is_alive())
    )


def create_blank_preview(select_text):
    """创建当前模板的空白可编辑页签，用于不选择文件的手工填写。"""
    global preview_select_text, preview_files, active_tree
    headers = get_core_headers(select_text)
    header_fields, detail_fields = get_preview_layout(select_text)
    file_result = {
        "filename": MANUAL_FILENAME,
        "status": MANUAL_STATUS,
        "message": "未选择文件，等待人工填写",
        "rows": [],
        "req_uuid": "",
        "log_row": [MANUAL_FILENAME, "", "", "", "", MANUAL_STATUS, MANUAL_STATUS, ""],
        "manual": True,
    }
    preview_select_text = select_text
    preview_files = [file_result]
    active_tree = None
    _build_file_tab(
        file_result, headers, header_fields, detail_fields, select_text
    )
    preview_notebook.select(0)
    on_preview_tab_changed()
    set_progress_state(0, "处理进度：待开始")


def on_template_changed(_event=None):
    """选择模板时清空旧预览并预渲染当前模板的空白页。"""
    global last_combo_text
    selected_text = combo_model.get().strip()
    if selected_text not in MODEL_MAP:
        return
    if _background_task_active():
        combo_model.set(last_combo_text)
        return
    last_combo_text = selected_text
    clear_preview()
    create_blank_preview(selected_text)


def _active_preview_file():
    """返回当前选中的预览文件数据，未选中时返回 None。"""
    selected_tab = preview_notebook.select()
    for info in preview_files:
        if str(info["tab"]) == selected_tab:
            return info
    return None


def show_preview(select_text, headers, file_results,
                 success_count, fail_count):
    """按文件创建预览页签，并调整界面按钮状态。"""
    global preview_select_text, preview_files, active_tree
    preview_select_text = select_text

    for info in preview_files:
        info["tab"].destroy()
    preview_files = list(file_results)
    active_tree = None
    header_fields, detail_fields = get_preview_layout(select_text)
    for file_result in preview_files:
        _build_file_tab(
            file_result, headers, header_fields, detail_fields, select_text
        )

    if preview_notebook.tabs():
        preview_notebook.select(0)
    on_preview_tab_changed()
    mock_check.config(state=tk.NORMAL)
    btn.config(text="选择文件并开始处理", state=tk.NORMAL)
    abort_btn.config(state=tk.DISABLED)
    total_rows = sum(len(info.get("rows") or []) for info in preview_files)
    pending_count = sum(
        info["status"] == "结果未生成" for info in preview_files
    )
    print_log(f"预览数据就绪：模板 {select_text}，页签数 {len(preview_files)}，"
              f"明细行数 {total_rows}，"
              f"成功 {success_count}，失败 {fail_count}，"
              f"结果未生成 {pending_count}")
    set_progress_state(100, "处理进度：完成")


def on_preview_tab_changed(_event=None):
    """切换页签时更新当前预览文件名和操作按钮状态。"""
    global active_tree
    info = _active_preview_file()
    if info is not None:
        if active_tree is not None and active_tree is not info["tree"]:
            active_tree.clear_clipboard()
        active_tree = info["tree"]
        current_file_label.config(text=f"当前预览文件：{info['filename']}")
        continue_btn.config(
            state=tk.NORMAL if not continue_query_active
            and info["status"] == "结果未生成" else tk.DISABLED
        )
        refresh_export_state()
        return
    if active_tree is not None:
        active_tree.clear_clipboard()
    active_tree = None
    current_file_label.config(text="当前预览文件：未选择")
    continue_btn.config(state=tk.DISABLED)
    refresh_export_state()


def refresh_row_action_state():
    """按当前选中行与复制内容刷新插入/复制/粘贴按钮。"""
    active = active_tree is not None
    active_info = _active_preview_file()
    editable = active and active_info is not None
    insert_btn.config(state=tk.NORMAL if editable else tk.DISABLED)
    copy_btn.config(
        state=tk.NORMAL if editable and active_tree.has_copyable_selection()
        else tk.DISABLED
    )
    paste_btn.config(
        state=tk.NORMAL if editable and active_tree.has_clipboard() else tk.DISABLED
    )


def refresh_export_state():
    """根据所有页签当前明细行数和操作状态刷新底部按钮。"""
    has_rows = any(bool(info["tree"].get_data()[1]) for info in preview_files)
    active_info = _active_preview_file()
    active_has_rows = active_info is not None and bool(
        active_info["tree"].get_data()[1]
    )
    tree_editable = active_tree is not None and active_info is not None
    refresh_row_action_state()
    export_btn.config(state=tk.NORMAL if has_rows else tk.DISABLED)
    wms_send_btn.config(
        state=tk.NORMAL
        if active_has_rows
        and preview_select_text in (
            "GE-发票单", "GE-ORACLE拣货单", "GE-OSCAR拣货单"
        )
        and not _background_task_active() and not wms_send_active
        else tk.DISABLED
    )
    add_btn.config(state=tk.NORMAL if tree_editable else tk.DISABLED)
    del_btn.config(state=tk.NORMAL if tree_editable else tk.DISABLED)


def active_tree_add_row():
    """在当前预览页签新增一行明细并刷新导出状态。"""
    if active_tree:
        active_tree.add_row()
        refresh_export_state()


def active_tree_delete_selected():
    """在当前预览页签删除选中行并刷新导出状态。"""
    if active_tree:
        active_tree.delete_selected()
        refresh_export_state()


def active_tree_insert_row():
    """在当前预览页签选中行下方插入空白行并刷新操作状态。"""
    if active_tree:
        active_tree.insert_row_after_selection()
        refresh_export_state()


def active_tree_copy_selected():
    """复制当前预览页签选中的整行并刷新操作状态。"""
    if active_tree:
        active_tree.copy_selected()
        refresh_export_state()


def active_tree_paste_row():
    """把当前预览页签复制的整行粘贴为新明细行并刷新操作状态。"""
    if active_tree:
        active_tree.paste_clipboard()
        refresh_export_state()


def clear_preview():
    """清空所有预览页签并恢复初始状态。"""
    global preview_select_text, preview_files, active_tree, continue_query_active
    for info in preview_files:
        info["tab"].destroy()
    preview_files = []
    active_tree = None
    continue_query_active = False
    preview_select_text = ""
    current_file_label.config(text="当前预览文件：未选择")
    continue_btn.config(state=tk.DISABLED)
    refresh_export_state()
    set_progress_state(0, "处理进度：待开始")
    btn.config(text="选择文件并开始处理", state=tk.NORMAL)
    abort_btn.config(state=tk.DISABLED)
    mock_check.config(state=tk.NORMAL)


def continue_current_task():
    """继续查询当前“结果未生成”页签记录的原 OCR 任务。"""
    global continue_query_active, continue_thread
    if continue_query_active:
        return
    info = _active_preview_file()
    if info is None or info.get("status") != "结果未生成":
        return
    req_uuid = str(info.get("req_uuid", "")).strip()
    if not req_uuid:
        messagebox.showwarning("温馨提示", "当前文件缺少原任务 reqUuid，无法继续查询")
        return

    abort_event.clear()
    continue_query_active = True
    continue_btn.config(state=tk.DISABLED)
    abort_btn.config(state=tk.NORMAL)
    set_progress_state(
        50,
        f"处理进度：正在续查 {os.path.basename(info['filename'])}...",
        "#D97706",
    )
    continue_thread = threading.Thread(
        target=continue_task_worker,
        args=(info, preview_select_text, req_uuid, abort_event),
        daemon=True,
    )
    continue_thread.start()
    refresh_export_state()
    win.after(100, poll_ui_queue)


def continue_task_worker(info, select_text, req_uuid, cancel_event):
    """后台继续查询原 OCR 任务，成功后解析并回传预览刷新消息。"""
    filename = info["filename"]
    print_log(f"继续查询原任务 reqUuid={req_uuid}，文件：{filename}")
    try:
        _, ocr_result_dict = call_get_result_api(req_uuid, cancel_event)
        if cancel_event.is_set():
            raise OCRAborted("OCR识别已由用户中止")
        if not (ocr_result_dict and ocr_result_dict.get("status") is True):
            raise Exception("OCR返回识别状态异常")
        commit_result = ocr_result_dict.get("data", {}).get("commitResult", {})
        parsed_rows, split_groups = parse_commit_result(
            select_text, commit_result, filename
        )
        updated_log_row = list(info.get("log_row", []))
        if len(updated_log_row) > 6:
            updated_log_row[5] = "成功"
            updated_log_row[6] = "处理成功"
        if cancel_event.is_set():
            raise OCRAborted("OCR识别已由用户中止")
        print_log(f"✅ [{filename}] 继续查询成功")
        ui_message_queue.put((
            "continue_success", info, select_text, "成功", "处理成功",
            parsed_rows, split_groups, updated_log_row,
        ))
    except OCRResultTimeout:
        print_log(f"⏳ [{filename}] 继续查询5分钟仍未生成结果")
        ui_message_queue.put((
            "continue_pending", info,
            "再次查询5分钟仍未生成结果，可继续查询原任务",
        ))
    except OCRAborted:
        print_log(f"⏹ [{filename}] 继续查询已由用户中止")
        ui_message_queue.put(("continue_aborted", info))
    except Exception as e:
        print_log(f"❌ [{filename}] 继续查询失败：{e}")
        ui_message_queue.put((
            "continue_pending", info, f"继续查询失败：{e}",
        ))


def _apply_continue_success(info, select_text):
    """按续查成功结果重建对应页签，并刷新可编辑/导出状态。"""
    headers = get_core_headers(select_text)
    header_fields, detail_fields = get_preview_layout(select_text)
    _replace_file_tab(info, headers, header_fields, detail_fields, select_text)
    refresh_export_state()


def _manual_export_base_name(select_text, header_values):
    """根据手工空白页所属模板取单据编号作为导出基准名。"""
    field_by_template = {
        "GE-ORACLE拣货单": "Order Number",
        "GE-OSCAR拣货单": "服务申请号",
        "GE-发票单": "INVOICE NO",
    }
    value = str(
        header_values.get(field_by_template.get(select_text, ""), "")
    ).strip()
    for char in ("\\", "/", ":", "*", "?", '"', "<", ">", "|"):
        value = value.replace(char, "_")
    return value or MANUAL_FILENAME


def start_export():
    """选择导出目录后收集有明细的文件页签并启动批量导出线程。"""
    missing_files = []
    for info in preview_files:
        _, detail_rows = info["tree"].get_data()
        if detail_rows and not str(
            info.get("header_values", {}).get("订单类型", "")
        ).strip():
            missing_files.append(info["filename"])
    if missing_files:
        messagebox.showwarning(
            "温馨提示",
            "以下文件请先选择订单类型：\n" + "\n".join(missing_files),
        )
        return

    missing_tracking_files = []
    if preview_select_text == "GE-发票单":
        for info in preview_files:
            _, detail_rows = info["tree"].get_data()
            if detail_rows and not str(
                info.get("header_values", {}).get("运单号", "")
            ).strip():
                missing_tracking_files.append(info["filename"])
    if missing_tracking_files:
        messagebox.showwarning(
            "温馨提示",
            "以下文件请先填写运单号：\n" + "\n".join(missing_tracking_files),
        )
        return

    export_targets = []
    for info in preview_files:
        _, detail_rows = info["tree"].get_data()
        if detail_rows:
            full_rows = merge_preview_rows(
                preview_select_text, info["header_values"], detail_rows
            )
            export_targets.append((info, full_rows))
    if not export_targets:
        messagebox.showwarning("温馨提示", "没有可导出的明细数据")
        return
    last_dir = get_last_export_dir()
    initial_dir = last_dir if last_dir and os.path.isdir(last_dir) else get_output_dir()
    output_dir = filedialog.askdirectory(
        title="选择Excel导出目录",
        initialdir=initial_dir,
    )
    if not output_dir:
        return
    save_last_export_dir(output_dir)
    export_btn.config(state=tk.DISABLED)
    global export_thread
    export_thread = threading.Thread(
        target=export_worker, args=(export_targets, output_dir), daemon=True
    )
    export_thread.start()
    refresh_export_state()
    win.after(100, poll_ui_queue)


def export_worker(export_targets, output_dir):
    """后台逐文件导出 Excel 到指定目录，汇总成功和失败消息。"""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    exported = []
    failures = []
    for info, rows in export_targets:
        if info.get("manual"):
            base_name = _manual_export_base_name(
                preview_select_text, info.get("header_values", {})
            )
        else:
            base_name = os.path.splitext(os.path.basename(info["filename"]))[0]
        output_file = os.path.join(
            output_dir, f"{base_name}_识别结果_{timestamp}.xlsx"
        )
        suffix = 2
        while os.path.exists(output_file):
            output_file = os.path.join(
                output_dir, f"{base_name}_识别结果_{timestamp}_{suffix}.xlsx"
            )
            suffix += 1
        try:
            export_excel(
                preview_select_text, rows, [info["log_row"]], output_file
            )
            exported.append(os.path.basename(output_file))
        except PermissionError:
            failures.append(f"{info['filename']}：文件被占用")
        except Exception as e:
            failures.append(f"{info['filename']}：{e}")

    if not exported and failures:
        ui_message_queue.put(("export_error", "Excel导出失败！\n" + "\n".join(failures)))
        return
    if failures:
        msg = f"导出完成！\n成功 {len(exported)} 个，失败 {len(failures)} 个。\n\n"
        msg += "\n".join(failures)
        ui_message_queue.put(("complete", msg))
        return
    msg = f"导出完成！\n已生成 {len(exported)} 个文件：\n" + "\n".join(exported)
    ui_message_queue.put(("complete", msg))


def close_wms_window():
    """关闭接口发送二级窗口并清理界面引用。"""
    global wms_window, wms_window_request_text, wms_window_response_text
    global wms_confirm_button
    if wms_window is not None:
        try:
            wms_window.destroy()
        except tk.TclError:
            pass
    wms_window = None
    wms_window_request_text = None
    wms_window_response_text = None
    wms_confirm_button = None


def _release_wms_grab_on_iconify(_event=None):
    """主窗口最小化时释放二级窗口抓取，避免任务栏恢复被模态状态阻塞。"""
    if wms_window is not None:
        try:
            wms_window.grab_release()
        except tk.TclError:
            pass


def _reapply_wms_grab():
    """重新给接口发送二级窗口设置模态抓取。"""
    if wms_window is None:
        return
    try:
        if wms_window.winfo_exists():
            wms_window.grab_set()
    except tk.TclError:
        pass


def _restore_wms_window_on_map(_event=None):
    """主窗口恢复时同步恢复接口发送二级窗口并重新建立抓取。"""
    if wms_window is None:
        return
    try:
        if not wms_window.winfo_exists():
            return
        wms_window.deiconify()
        wms_window.lift()
        wms_window.focus_force()
        win.after_idle(_reapply_wms_grab)
    except tk.TclError:
        pass


def open_wms_send_window():
    """打开当前页签的只读报文窗口，支持确认发送和回告展示。"""
    global wms_window, wms_window_request_text, wms_window_response_text
    global wms_confirm_button, wms_thread
    global wms_send_active
    if wms_send_active or _background_task_active():
        return
    if preview_select_text not in (
        "GE-发票单", "GE-ORACLE拣货单", "GE-OSCAR拣货单"
    ):
        return
    info = _active_preview_file()
    if info is None:
        return
    _, detail_rows = info["tree"].get_data()
    if not detail_rows:
        messagebox.showwarning("温馨提示", "当前单据没有可发送的明细数据")
        return
    header_values = info.get("header_values", {})
    if not str(header_values.get("订单类型", "")).strip():
        messagebox.showwarning("温馨提示", "当前单据请先选择订单类型")
        return
    if preview_select_text == "GE-发票单":
        if not str(header_values.get("运单号", "")).strip():
            messagebox.showwarning("温馨提示", "当前发票缺少运单号，无法发送")
            return
        if not str(header_values.get("INVOICE NO", "")).strip():
            messagebox.showwarning("温馨提示", "当前发票缺少INVOICE NO，无法发送")
            return
        payload = build_put_purchase_order_payload(header_values, detail_rows)
        send_func = send_put_purchase_order
        log_name = "采购订单"
    elif preview_select_text == "GE-ORACLE拣货单":
        if not str(header_values.get("Order Number", "")).strip():
            messagebox.showwarning("温馨提示", "当前ORACLE拣货单缺少Order Number，无法发送")
            return
        payload = build_put_original_sales_order_payload(
            preview_select_text, header_values, detail_rows
        )
        send_func = send_put_original_sales_order
        log_name = "ORACLE销售订单"
    else:
        if not str(header_values.get("服务申请号", "")).strip():
            messagebox.showwarning("温馨提示", "当前OSCAR拣货单缺少服务申请号，无法发送")
            return
        payload = build_put_original_sales_order_payload(
            preview_select_text, header_values, detail_rows
        )
        send_func = send_put_original_sales_order
        log_name = "OSCAR销售订单"
    if wms_window is not None:
        try:
            if wms_window.winfo_exists():
                wms_window.lift()
                return
        except tk.TclError:
            pass

    wms_window = tk.Toplevel(win)
    wms_window.title("接口发送")
    wms_window.geometry("1000x720")
    wms_window.minsize(720, 480)
    wms_window.transient(win)
    wms_window.grab_set()
    wms_window.protocol("WM_DELETE_WINDOW", close_wms_window)

    body = tk.Frame(wms_window)
    body.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    button_frame = tk.Frame(body)
    button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

    paned = tk.PanedWindow(body, orient=tk.VERTICAL, sashwidth=6)
    paned.pack(fill=tk.BOTH, expand=True)

    request_frame = tk.Frame(paned)
    response_frame = tk.Frame(paned)
    paned.add(request_frame, minsize=160)
    paned.add(response_frame, minsize=120)

    wms_window_request_text = _wms_text_pane(request_frame, "组装报文")
    wms_window_response_text = _wms_text_pane(response_frame, "接口返回内容")

    wms_window_request_text.config(state=tk.NORMAL)
    wms_window_request_text.insert(
        tk.END, json.dumps(payload, ensure_ascii=False, indent=2)
    )
    wms_window_request_text.config(state=tk.DISABLED)

    wms_window_response_text.config(state=tk.NORMAL)
    wms_window_response_text.insert(tk.END, "尚未发送")
    wms_window_response_text.config(state=tk.DISABLED)

    def start_wms_send():
        global wms_thread, wms_send_active
        if wms_send_active or wms_window_response_text is None:
            return
        wms_send_active = True
        if wms_confirm_button is not None:
            wms_confirm_button.config(state=tk.DISABLED, text="发送中...")
        token = id(wms_window_response_text)
        wms_thread = threading.Thread(
            target=wms_send_worker,
            args=(payload, token, send_func, log_name),
            daemon=True,
        )
        wms_thread.start()
        refresh_export_state()
        win.after(100, poll_ui_queue)

    wms_confirm_button = tk.Button(
        button_frame, text="确认发送", command=start_wms_send,
        width=12, bg="#0E7490", fg="#111827", font=BUTTON_FONT,
        activebackground="#155E75", activeforeground="#111827",
        disabledforeground=DISABLED_FOREGROUND,
    )
    tk.Button(
        button_frame, text="关闭", command=close_wms_window,
        width=10, font=BUTTON_FONT,
        disabledforeground=DISABLED_FOREGROUND,
    ).pack(side=tk.RIGHT)
    wms_confirm_button.pack(side=tk.RIGHT, padx=(0, 8))

    wms_window.wait_visibility()
    paned.update_idletasks()
    try:
        paned.sash_place(0, 0, max(160, int(paned.winfo_height() * 0.7)))
    except tk.TclError:
        pass

    wms_window_request_text.yview_moveto(0)
    wms_window_response_text.yview_moveto(0)


def wms_send_worker(payload, token, send_func, log_name):
    """后台发送 WMS 报文，并把回告文本回传主线程。"""
    print_log(f"正在发送WMS{log_name}报文...")
    try:
        response = send_func(payload)
        text = format_wms_response(response)
        print_log(f"WMS接口回告：{text[:200]}")
        if is_wms_send_success(response):
            result_text = f"发送成功\n\n{text}"
            ui_message_queue.put(("wms_send_result", token, result_text))
        else:
            result_text = f"发送失败：HTTP 状态或 returnFlag 不满足\n\n{text}"
            ui_message_queue.put(("wms_send_result", token, result_text))
    except Exception as e:
        print_log(f"WMS接口发送失败: {e}")
        error_text = f"发送失败：{e}"
        ui_message_queue.put(("wms_send_result", token, error_text))


def draw_progress_canvas():
    """按当前进度百分比重绘进度条。"""
    width = progress_canvas.winfo_width()
    height = progress_canvas.winfo_height()
    if width <= 1 or height <= 1:
        return
    progress_canvas.delete("all")
    progress_canvas.create_rectangle(
        1, 1, width - 1, height - 1, fill="#E5E7EB", outline="#94A3B8"
    )
    fill_width = int((width - 2) * progress_percent / 100)
    if fill_width > 0:
        progress_canvas.create_rectangle(
            1, 1, fill_width + 1, height - 1, fill=progress_color, outline=""
        )


def set_progress_state(percent, text, color="#16A34A"):
    """更新进度文案、颜色和进度条。"""
    global progress_percent, progress_color
    progress_percent = percent
    progress_color = color
    progress_label.config(
        text=text,
        fg="#B42318" if color == "#D92D20" else "#111827"
    )
    draw_progress_canvas()


def update_progress(done, total):
    percent = 50 + round(done / total * 50) if total else 100
    set_progress_state(percent, f"处理进度：{done}/{total}（{percent}%）")


def poll_ui_queue():
    """主线程轮询处理结果消息，并驱动界面状态更新。"""
    global continue_query_active, wms_send_active
    flush_log()
    while True:
        try:
            kind, *payload = ui_message_queue.get_nowait()
        except queue.Empty:
            break
        if kind == "preview":
            if abort_event.is_set():
                finish_abort_state()
                continue
            show_preview(*payload)
        elif kind == "progress":
            update_progress(payload[0], payload[1])
        elif kind == "complete":
            refresh_export_state()
            messagebox.showinfo("完成", payload[0])
        elif kind == "processing_aborted":
            finish_abort_state()
        elif kind == "processing_error":
            btn.config(text="选择文件并开始处理", state=tk.NORMAL)
            mock_check.config(state=tk.NORMAL)
            abort_btn.config(state=tk.DISABLED)
            set_progress_state(100, "处理进度：处理失败", "#D92D20")
            messagebox.showerror("错误", payload[0])
        elif kind == "export_error":
            refresh_export_state()
            messagebox.showerror("错误", payload[0])
        elif kind == "continue_aborted":
            info = payload[0]
            info["message"] = "继续查询已中止"
            continue_query_active = False
            abort_event.clear()
            _refresh_file_status_label(info)
            on_preview_tab_changed()
            finish_abort_state()
        elif kind == "continue_success":
            (info, select_text, status, message, parsed_rows,
             split_groups, updated_log_row) = payload
            info["status"] = status
            info["message"] = message
            info["rows"] = parsed_rows
            info["split_groups"] = split_groups
            info["log_row"] = updated_log_row
            continue_query_active = False
            _apply_continue_success(info, select_text)
            set_progress_state(100, "处理进度：续查完成")
        elif kind == "continue_pending":
            info, message = payload
            info["message"] = message
            continue_query_active = False
            _refresh_file_status_label(info)
            on_preview_tab_changed()
            set_progress_state(100, "处理进度：续查未生成结果", "#D97706")
        elif kind == "wms_send_result":
            token, text = payload
            _replace_wms_response(token, text)
            wms_send_active = False
            if wms_confirm_button is not None:
                try:
                    if wms_confirm_button.winfo_exists():
                        wms_confirm_button.config(
                            state=tk.NORMAL, text="重新发送"
                        )
                except tk.TclError:
                    pass
            refresh_export_state()

    if any(
        thread is not None and thread.is_alive()
        for thread in (worker_thread, export_thread, continue_thread, wms_thread)
    ):
        win.after(100, poll_ui_queue)


# ========== 界面部分 ==========
win = tk.Tk()
win.title("GE单据批量OCR处理工具")
max_width, max_height = win.maxsize()
win.geometry(f"{max_width}x{max_height}")
win.bind("<Unmap>", _release_wms_grab_on_iconify)
win.bind("<Map>", _restore_wms_window_on_map)

top = tk.Frame(win)
top.pack(fill=tk.X, padx=10, pady=(10, 0))

BUTTON_FONT = tkfont.nametofont("TkDefaultFont").copy()
BUTTON_FONT.configure(weight="bold")
DISABLED_FOREGROUND = "#111827"

tk.Label(top, text="选择模版规则：", font=("黑体", 11)).pack(side=tk.LEFT)
combo_model = ttk.Combobox(top, width=28, font=("黑体", 11), state="readonly")
combo_model["values"] = list(MODEL_MAP.keys())
combo_model.set("")
combo_model.pack(side=tk.LEFT, padx=(0, 20))
last_combo_text = combo_model.get()
combo_model.bind("<<ComboboxSelected>>", on_template_changed)

mock_var = tk.BooleanVar(value=False)
mock_check = tk.Checkbutton(top, text="模拟数据", variable=mock_var, font=("黑体", 11))
mock_check.pack(side=tk.LEFT, padx=(0, 20))

btn = tk.Button(top, text="选择文件并开始处理", command=run_task, width=22,
                bg="#4CAF50", fg="#0B3D0F", font=BUTTON_FONT,
                disabledforeground=DISABLED_FOREGROUND)
btn.pack(side=tk.LEFT)
export_btn = tk.Button(top, text="确认并导出", command=start_export,
                       width=16, bg="#2196F3", fg="#0A2540", font=BUTTON_FONT,
                       disabledforeground=DISABLED_FOREGROUND,
                       state=tk.DISABLED)
export_btn.pack(side=tk.LEFT, padx=(8, 0))
wms_send_btn = tk.Button(
    top, text="接口发送", command=open_wms_send_window,
    width=10, bg="#0E7490", fg="#111827", font=BUTTON_FONT,
    activebackground="#155E75", activeforeground="#111827",
    disabledforeground=DISABLED_FOREGROUND,
    state=tk.DISABLED,
)
wms_send_btn.pack(side=tk.LEFT, padx=(8, 0))
abort_btn = tk.Button(
    top, text="中止", command=abort_processing, width=10,
    bg="#D92D20", fg="#111827", font=BUTTON_FONT,
    activebackground="#B42318", activeforeground="#111827",
    disabledforeground=DISABLED_FOREGROUND,
    state=tk.DISABLED,
)
abort_btn.pack(side=tk.LEFT, padx=(8, 0))

progress_frame = tk.Frame(top)
progress_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

progress_label = tk.Label(progress_frame, text="处理进度：待开始",
                          font=("黑体", 11, "bold"), anchor="w")
progress_label.pack(side=tk.LEFT)

progress_canvas = tk.Canvas(progress_frame, height=14, highlightthickness=0)
progress_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0), pady=7)
progress_canvas.bind("<Configure>", lambda _event: draw_progress_canvas())

table_frame = tk.Frame(win)
table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(2, 10))

current_file_label = tk.Label(
    table_frame, text="当前预览文件：未选择",
    font=("黑体", 11, "bold"), anchor="w"
)
current_file_label.pack(fill=tk.X, padx=2, pady=(0, 4))

preview_notebook = ttk.Notebook(table_frame)
preview_notebook.pack(fill=tk.BOTH, expand=True)
preview_notebook.bind("<<NotebookTabChanged>>", on_preview_tab_changed)

style = ttk.Style(win)
if sys.platform == "win32":
    style.theme_use("clam")
style.configure("Preview.Treeview", background="#FFFFFF",
                fieldbackground="#FFFFFF", rowheight=39)
style.configure("Preview.Treeview.Heading", background="#E8EEF2",
                font=PREVIEW_HEADING_FONT, anchor="center")

op_frame = tk.Frame(win)
op_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

add_btn = tk.Button(op_frame, text="新增行", command=active_tree_add_row,
                    width=10, font=BUTTON_FONT,
                    disabledforeground=DISABLED_FOREGROUND,
                    state=tk.DISABLED)
add_btn.pack(side=tk.LEFT, padx=(0, 8))
insert_btn = tk.Button(op_frame, text="插入行", command=active_tree_insert_row,
                       width=10, font=BUTTON_FONT,
                       disabledforeground=DISABLED_FOREGROUND,
                       state=tk.DISABLED)
insert_btn.pack(side=tk.LEFT, padx=(0, 8))
copy_btn = tk.Button(op_frame, text="复制行", command=active_tree_copy_selected,
                     width=10, font=BUTTON_FONT,
                     disabledforeground=DISABLED_FOREGROUND,
                     state=tk.DISABLED)
copy_btn.pack(side=tk.LEFT, padx=(0, 8))
paste_btn = tk.Button(op_frame, text="粘贴行", command=active_tree_paste_row,
                      width=10, font=BUTTON_FONT,
                      disabledforeground=DISABLED_FOREGROUND,
                      state=tk.DISABLED)
paste_btn.pack(side=tk.LEFT, padx=(0, 8))
del_btn = tk.Button(op_frame, text="删除行", command=active_tree_delete_selected,
                    width=10, font=BUTTON_FONT,
                    disabledforeground=DISABLED_FOREGROUND,
                    state=tk.DISABLED)
del_btn.pack(side=tk.LEFT, padx=(0, 8))
continue_btn = tk.Button(
    op_frame, text="继续查询原任务", command=continue_current_task,
    width=16, bg="#D97706", fg="#111827", font=BUTTON_FONT,
    disabledforeground=DISABLED_FOREGROUND,
    state=tk.DISABLED,
)
continue_btn.pack(side=tk.LEFT, padx=(10, 0))

query_log_btn = tk.Button(
    op_frame, text="查询日志", command=open_log_window,
    width=10, font=BUTTON_FONT,
    disabledforeground=DISABLED_FOREGROUND,
)
query_log_btn.pack(side=tk.LEFT, padx=(8, 0))

win.after(200, poll_log_queue)

if __name__ == "__main__":
    win.state("zoomed")
    win.mainloop()
