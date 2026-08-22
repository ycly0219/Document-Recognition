"""Excel 识别结果与处理日志导出。"""

import os
import re
import time
from copy import copy
from datetime import date

from openpyxl import Workbook, load_workbook

from logging_utils import print_log
from parsers import get_core_headers

PO_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "DOC_PO_HEADER.xlsx"
)

_DATE_PATTERN = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$")

_INVOICE_FIXED_VALUES = {
    "A": "WH004078",
    "B": "OSI",
    "C": "00",
    "D": "GEHC",
    "R": "WH004078",
    "S": "GEHC",
    "V": "EA",
    "AC": "GOOD",
}

_INVOICE_DYNAMIC_COLUMNS = {
    "F": 0,   # INVOICE NO
    "L": 12,  # CARRIER
    "M": 13,  # HAWB
    "T": 1,   # ITEM NUMBER
    "U": 2,   # QTY
    "X": 6,   # Expiration Date
    "Z": 5,   # LOT Number
    "AD": 4,  # Serial Number
    "AF": 3,  # LPN Number
    "AJ": 7,  # COUNTRY OF ORIGIN
    "AK": 8,  # SALES ORDER NO
    "AL": 9,  # CUSTOMER PO
}


def _row_value(row, index):
    """读取预览行字段，兼容界面手工新增的不完整行。"""
    return row[index] if index < len(row) else ""


def _normalize_expiration_date(value):
    """把 YYYY/MM/DD 或 YYYY-MM-DD 转成模板要求的 YYYY-MM-DD。"""
    if value is None:
        return ""
    text = str(value).strip()
    match = _DATE_PATTERN.fullmatch(text)
    if not match:
        return text
    try:
        year, month, day = (int(part) for part in match.groups())
        return date(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return text


def _export_invoice_po_template(core_data, output_file):
    """按 DOC_PO_HEADER 的采购订单表头模板导出发票数据。"""
    wb = load_workbook(PO_TEMPLATE_PATH)
    ws = wb["采购订单表头"]
    del wb["系统代码说明"]

    now = time.localtime()
    creation_time = f"{now.tm_year}/{now.tm_mon}/{now.tm_mday} " \
                    f"{time.strftime('%H:%M:%S', now)}"

    for row_index, preview_row in enumerate(core_data, start=3):
        if row_index > 3:
            for col_index in range(1, ws.max_column + 1):
                template_cell = ws.cell(row=3, column=col_index)
                target_cell = ws.cell(row=row_index, column=col_index)
                target_cell._style = copy(template_cell._style)

        ws[f"E{row_index}"] = creation_time
        for col_name, value in _INVOICE_FIXED_VALUES.items():
            ws[f"{col_name}{row_index}"] = value
        for col_name, source_index in _INVOICE_DYNAMIC_COLUMNS.items():
            value = _row_value(preview_row, source_index)
            if col_name == "X":
                value = _normalize_expiration_date(value)
            ws[f"{col_name}{row_index}"] = value

    wb.save(output_file)


def export_excel(select_text, core_data, full_log_data, output_file=None):
    """生成单个文件的识别结果 Excel，返回输出文件路径。"""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_file = output_file or os.path.join(
        os.getcwd(), f"GE单据OCR识别结果_{timestamp}.xlsx"
    )

    if select_text == "GE-发票单":
        _export_invoice_po_template(core_data, output_file)
        print_log(f"Excel导出完成，路径：{output_file}")
        return output_file

    wb = Workbook()
    ws_core = wb.active
    ws_core.title = "识别结果列表"

    core_headers = get_core_headers(select_text)

    ws_core.append(core_headers)
    for row in core_data:
        ws_core.append(row)

    # 自适应列宽
    for col in ws_core.columns:
        max_len = max(len(str(cell.value)) for cell in col)
        ws_core.column_dimensions[col[0].column_letter].width = min(max_len + 2, 80)

    # 日志Sheet
    ws_log = wb.create_sheet(title="处理日志")
    log_headers = ["文件名", "fileId", "文件URL", "reqUuid",
                   "提取单据编号", "状态", "异常信息", "识别报文"]
    ws_log.append(log_headers)
    for row in full_log_data:
        ws_log.append(row)
    for col in ws_log.columns:
        max_len = max(len(str(cell.value)) for cell in col)
        ws_log.column_dimensions[col[0].column_letter].width = min(max_len + 2, 80)

    wb.save(output_file)
    print_log(f"Excel导出完成，路径：{output_file}")
    return output_file
