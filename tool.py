"""Tkinter 入口模块，负责界面、后台线程与批量处理编排。"""

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
    call_get_result_api,
    call_process_api,
    upload_file_to_server,
)
from parsers import get_core_headers, parse_commit_result


ui_message_queue = queue.Queue()
worker_thread = None
export_thread = None
abort_event = threading.Event()
preview_select_text = ""
preview_files = []
active_tree = None
progress_percent = 0
progress_color = "#16A34A"

if sys.platform == "win32":
    PREVIEW_HEADING_FONT = ("Microsoft YaHei UI", 10, "bold")
else:
    PREVIEW_HEADING_FONT = ("黑体", 15, "bold")
MAX_BATCH_FILES = 5


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

    def _selected_rows_in_order(self):
        rows = self.selection()
        if not rows:
            return []
        children = self.get_children()
        position = {row_id: index for index, row_id in enumerate(children)}
        return sorted(rows, key=lambda row_id: position.get(row_id, len(children)))

    def has_clipboard(self):
        return bool(self._clipboard)

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
            children = self.get_children()
            index = children.index(selected[-1]) + 1
        else:
            index = tk.END
        row_id = self.insert(
            "", index, values=("",) * len(self["columns"]),
            tags=("new_row",)
        )
        self._renumber()
        self.selection_set(row_id)
        self.see(row_id)
        return row_id

    def paste_clipboard(self):
        if not self._clipboard or not self["columns"]:
            return False
        selected = self._selected_rows_in_order()
        children = self.get_children()
        if selected:
            index = children.index(selected[-1]) + 1
        else:
            index = len(children)
        pasted = []
        for offset, values in enumerate(self._clipboard):
            row_id = self.insert(
                "", index + offset, values=list(values),
                tags=("new_row",)
            )
            pasted.append(row_id)
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
            self.delete(row_id)
        self._renumber()

    def _renumber(self):
        for index, row_id in enumerate(self.get_children(), 1):
            self.item(row_id, text=index)

    def auto_size_columns(self, max_width=320, min_width=90, padding=26):
        """按内容自动收缩列宽并居中，长内容由用户拖动列宽查看。"""
        font = tkfont.nametofont("TkDefaultFont")
        heading_font = tkfont.Font(font=PREVIEW_HEADING_FONT)
        for col in self["columns"]:
            content_w = [font.measure(self.set(row, col))
                         for row in self.get_children()]
            width = max([heading_font.measure(col)] + content_w) + padding
            width = min(max(min_width, width), max_width)
            self.column(col, width=width, minwidth=min_width,
                        stretch=False, anchor="center")
            self.heading(col, anchor="center")

    def clear_rows(self):
        for row_id in self.get_children():
            self.delete(row_id)

    def get_data(self):
        headers = list(self["columns"])
        rows = [[self.set(row_id, col) for col in headers]
                for row_id in self.get_children()]
        return headers, rows


def flush_log():
    """由主线程把日志队列写入文本框，避免跨线程操作 Tkinter。"""
    while True:
        try:
            line = log_queue.get_nowait()
        except queue.Empty:
            break
        log_text.insert(tk.END, line + "\n")
    log_text.see(tk.END)
    log_text.update_idletasks()


# ---------------- 主流程函数 ----------------
def run_task():
    """根据当前模板启动真实或模拟批量处理。"""
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
    """请求中止当前 OCR 处理批次。"""
    abort_event.set()
    abort_btn.config(state=tk.DISABLED)
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
                parsed_rows = parse_commit_result(select_text, commit_result, filename)
                core_data.extend(parsed_rows)
                if select_text == "GE-发票单":
                    print_log(f"{filename} 处理完成后总明细行数:{len(core_data)}")

                status = "成功"
                res_msg = "处理成功"
                success_count += 1
                print_log(f"✅ [{filename}] 处理成功")

            else:
                raise Exception("OCR返回识别状态异常")

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
            "log_row": [filename, file_id, file_url, req_uuid, "", status, res_msg, ""],
        })
        ui_message_queue.put(("progress", done, total_files))

    print_log(f"\n===== 批量处理结束 =====")
    print_log(f"总文件：{total_files} | 成功：{success_count} | 失败：{fail_count} | 有效明细行数：{len(core_data)}")
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


def _build_file_tab(file_result, headers):
    """为单个文件创建预览页签，并返回该页签对应的可编辑表格。"""
    tab = ttk.Frame(preview_notebook)
    preview_notebook.add(tab, text=_short_tab_label(file_result["filename"]))

    status_text = f"状态：{file_result['status']}"
    if not file_result.get("rows"):
        status_text += "；当前无明细数据"
    if file_result.get("message"):
        status_text += f"；{file_result['message']}"
    status_label = tk.Label(
        tab, text=status_text, anchor="w",
        fg="#B42318" if file_result["status"] == "失败" else "#111827"
    )
    status_label.pack(fill=tk.X, padx=8, pady=(0, 4))

    tree_frame = tk.Frame(tab)
    tree_frame.pack(fill=tk.BOTH, expand=True)
    tree = EditableTreeview(tree_frame, style="Preview.Treeview",
                            selectmode="extended")
    tree.bind("<<TreeviewSelect>>",
              lambda _event: refresh_row_action_state())
    tree.tag_configure("new_row", background="#FFF3CD")
    tree.tag_configure("zebra_even", background="#FFFFFF")
    tree.tag_configure("zebra_odd", background="#EFF5F9")
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tree_frame.rowconfigure(0, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    tree["columns"] = headers
    tree.column("#0", width=40, minwidth=30, stretch=False, anchor="center")
    for header in headers:
        tree.heading(header, text=header)
    for row_index, row in enumerate(file_result.get("rows") or []):
        tag = "zebra_even" if row_index % 2 == 0 else "zebra_odd"
        tree.insert("", tk.END, values=row, tags=(tag,))
    tree._renumber()
    tree.auto_size_columns()

    file_result["tree"] = tree
    file_result["tab"] = tab
    return tree


def show_preview(select_text, headers, file_results,
                 success_count, fail_count):
    """按文件创建预览页签，并调整界面按钮状态。"""
    global preview_select_text, preview_files, active_tree
    preview_select_text = select_text

    for info in preview_files:
        info["tab"].destroy()
    preview_files = list(file_results)
    active_tree = None
    for file_result in preview_files:
        _build_file_tab(file_result, headers)

    if preview_notebook.tabs():
        preview_notebook.select(0)
    on_preview_tab_changed()
    mock_check.config(state=tk.NORMAL)
    btn.config(text="选择文件并开始处理", state=tk.NORMAL)
    abort_btn.config(state=tk.DISABLED)
    total_rows = sum(len(info["tree"].get_children()) for info in preview_files)
    print_log(f"预览数据就绪：模板 {select_text}，页签数 {len(preview_files)}，"
              f"明细行数 {total_rows}，"
              f"成功 {success_count}，失败 {fail_count}")
    set_progress_state(100, "处理进度：完成")


def on_preview_tab_changed(_event=None):
    """切换页签时更新当前预览文件名和操作按钮状态。"""
    global active_tree
    selected_tab = preview_notebook.select()
    for info in preview_files:
        if str(info["tab"]) == selected_tab:
            if active_tree is not None and active_tree is not info["tree"]:
                active_tree.clear_clipboard()
            active_tree = info["tree"]
            current_file_label.config(text=f"当前预览文件：{info['filename']}")
            refresh_export_state()
            return
    if active_tree is not None:
        active_tree.clear_clipboard()
    active_tree = None
    current_file_label.config(text="当前预览文件：未选择")
    refresh_export_state()


def refresh_row_action_state():
    """按当前选中行与复制内容刷新插入/复制/粘贴按钮。"""
    active = active_tree is not None
    insert_btn.config(state=tk.NORMAL if active else tk.DISABLED)
    copy_btn.config(
        state=tk.NORMAL if active and active_tree.selection() else tk.DISABLED
    )
    paste_btn.config(
        state=tk.NORMAL if active and active_tree.has_clipboard() else tk.DISABLED
    )


def refresh_export_state():
    """根据所有页签当前明细行数和操作状态刷新底部按钮。"""
    has_rows = any(bool(info["tree"].get_children()) for info in preview_files)
    refresh_row_action_state()
    export_btn.config(state=tk.NORMAL if has_rows else tk.DISABLED)
    add_btn.config(state=tk.NORMAL if active_tree else tk.DISABLED)
    del_btn.config(state=tk.NORMAL if active_tree else tk.DISABLED)


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
    global preview_select_text, preview_files, active_tree
    for info in preview_files:
        info["tab"].destroy()
    preview_files = []
    active_tree = None
    preview_select_text = ""
    current_file_label.config(text="当前预览文件：未选择")
    refresh_export_state()
    set_progress_state(0, "处理进度：待开始")
    btn.config(text="选择文件并开始处理", state=tk.NORMAL)
    abort_btn.config(state=tk.DISABLED)
    mock_check.config(state=tk.NORMAL)


def start_export():
    """选择导出目录后收集有明细的文件页签并启动批量导出线程。"""
    export_targets = []
    for info in preview_files:
        _, rows = info["tree"].get_data()
        if rows:
            export_targets.append((info, rows))
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
    win.after(100, poll_ui_queue)


def export_worker(export_targets, output_dir):
    """后台逐文件导出 Excel 到指定目录，汇总成功和失败消息。"""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    exported = []
    failures = []
    for info, rows in export_targets:
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

    if (worker_thread is not None and worker_thread.is_alive()) or \
       (export_thread is not None and export_thread.is_alive()):
        win.after(100, poll_ui_queue)


# ========== 界面部分 ==========
win = tk.Tk()
win.title("GE单据批量OCR处理工具")
win.geometry("1180x760")

top = tk.Frame(win)
top.pack(fill=tk.X, padx=10, pady=(10, 0))

tk.Label(top, text="选择模版规则：", font=("黑体", 11)).pack(side=tk.LEFT)
combo_model = ttk.Combobox(top, width=28, font=("黑体", 11), state="readonly")
combo_model["values"] = list(MODEL_MAP.keys())
combo_model.set("")
combo_model.pack(side=tk.LEFT, padx=(0, 20))

mock_var = tk.BooleanVar(value=False)
mock_check = tk.Checkbutton(top, text="模拟数据", variable=mock_var, font=("黑体", 11))
mock_check.pack(side=tk.LEFT, padx=(0, 20))

btn = tk.Button(top, text="选择文件并开始处理", command=run_task, width=22,
                bg="#4CAF50", fg="#0B3D0F")
btn.pack(side=tk.LEFT)
export_btn = tk.Button(top, text="确认并导出", command=start_export,
                       width=14, bg="#2196F3", fg="#0A2540", state=tk.DISABLED)
export_btn.pack(side=tk.LEFT, padx=(8, 0))
abort_btn = tk.Button(
    top, text="中止", command=abort_processing, width=10,
    bg="#D92D20", fg="#111827",
    activebackground="#B42318", activeforeground="#111827",
    disabledforeground="#6B7280",
    state=tk.DISABLED,
)
abort_btn.pack(side=tk.LEFT, padx=(8, 0))

progress_frame = tk.Frame(win)
progress_frame.pack(fill=tk.X, padx=10, pady=(8, 0))

progress_label = tk.Label(progress_frame, text="处理进度：待开始",
                          font=("黑体", 11, "bold"), anchor="w")
progress_label.pack(fill=tk.X)

progress_canvas = tk.Canvas(progress_frame, height=28, highlightthickness=0)
progress_canvas.pack(fill=tk.X, pady=(4, 0))
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
                    width=10, state=tk.DISABLED)
add_btn.pack(side=tk.LEFT, padx=(0, 8))
insert_btn = tk.Button(op_frame, text="插入行", command=active_tree_insert_row,
                       width=10, state=tk.DISABLED)
insert_btn.pack(side=tk.LEFT, padx=(0, 8))
copy_btn = tk.Button(op_frame, text="复制行", command=active_tree_copy_selected,
                     width=10, state=tk.DISABLED)
copy_btn.pack(side=tk.LEFT, padx=(0, 8))
paste_btn = tk.Button(op_frame, text="粘贴行", command=active_tree_paste_row,
                      width=10, state=tk.DISABLED)
paste_btn.pack(side=tk.LEFT, padx=(0, 8))
del_btn = tk.Button(op_frame, text="删除行", command=active_tree_delete_selected,
                    width=10, state=tk.DISABLED)
del_btn.pack(side=tk.LEFT, padx=(0, 8))

tk.Label(win, text="处理日志", font=("黑体", 10)).pack(anchor="w", padx=12)
log_text = tk.Text(win, height=8, width=110)
log_text.pack(fill=tk.X, padx=10, pady=(2, 10))

if __name__ == "__main__":
    win.mainloop()
