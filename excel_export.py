"""Excel 识别结果与处理日志导出。"""

import os
import time

from openpyxl import Workbook

from logging_utils import print_log
from parsers import get_core_headers


def export_excel(select_text, core_data, full_log_data, output_file=None):
    """生成单个文件的识别结果 Excel，返回输出文件路径。"""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_file = output_file or os.path.join(
        os.getcwd(), f"GE单据OCR识别结果_{timestamp}.xlsx"
    )

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
