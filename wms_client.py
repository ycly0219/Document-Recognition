"""Flux WMS putPurchaseOrder 报文构建与发送。"""

import datetime
import json

import requests

from config import (
    WMS_CUSTOMER_ID,
    WMS_PUT_PURCHASE_ORDER_URL,
    WMS_WAREHOUSE_ID,
)
from excel_export import _normalize_expiration_date
from parsers import get_order_type_value


_WMS_DETAIL_HEADERS = [
    "ITEM NUMBER",
    "QTY",
    "LPN Number",
    "Serial Number",
    "LOT Number",
    "Expiration Date",
    "COUNTRY OF ORIGIN",
    "SALES ORDER NO",
    "CUSTOMER PO",
]


def _text(value):
    """把界面值统一转成 WMS 报文使用的字符串。"""
    return "" if value is None else str(value).strip()


def build_put_purchase_order_payload(header_values, detail_rows):
    """按发票页签的单据头和可导出发明细构建 putPurchaseOrder 报文。"""
    header = {
        "warehouseId": WMS_WAREHOUSE_ID,
        "customerId": WMS_CUSTOMER_ID,
        "poType": get_order_type_value(
            "GE-发票单", header_values.get("订单类型", "")
        ),
        "docNo": _text(header_values.get("INVOICE NO", "")),
        "udf01": _text(header_values.get("CARRIER", "")),
        "udf02": _text(header_values.get("HAWB", "")),
    }

    details = []
    for line_no, row in enumerate(detail_rows or [], start=1):
        row_map = dict(zip(_WMS_DETAIL_HEADERS, row))
        details.append({
            "lineNo": str(line_no),
            "customerId": WMS_CUSTOMER_ID,
            "sku": _text(row_map.get("ITEM NUMBER")),
            "orderedQty": _text(row_map.get("QTY")),
            "packUom": "EA",
            "lotAtt02": _normalize_expiration_date(
                row_map.get("Expiration Date")
            ),
            "lotAtt03": datetime.now().strftime("%Y-%m-%d"),
            "lotAtt04": _text(row_map.get("LOT Number")),
            "lotAtt05": "ORACLE",
            "lotAtt08": "GOOD",
            "lotAtt09": _text(row_map.get("Serial Number")),
            "lotAtt11": _text(row_map.get("LPN Number")),
            "lotAtt15": _text(row_map.get("COUNTRY OF ORIGIN")),
            "lotAtt16": _text(row_map.get("SALES ORDER NO")),
            "lotAtt17": _text(row_map.get("CUSTOMER PO")),
        })

    return {"data": {"header": [{**header, "details": details}]}}


def send_put_purchase_order(payload):
    """发送 putPurchaseOrder 报文并返回接口响应对象。"""
    return requests.post(
        WMS_PUT_PURCHASE_ORDER_URL, json=payload, timeout=30
    )


def is_wms_send_success(response):
    """按 HTTP 状态与顶层 returnFlag 判定 putPurchaseOrder 是否发送成功。"""
    if getattr(response, "status_code", None) != 200:
        return False
    try:
        body = response.json()
        return_flag = body["Response"]["return"]["returnFlag"]
    except (ValueError, AttributeError, KeyError, TypeError):
        return False
    return return_flag == "1" or (
        isinstance(return_flag, int)
        and not isinstance(return_flag, bool)
        and return_flag == 1
    )


def format_wms_response(response):
    """把接口回告整理为可读文本，优先展示格式化 JSON。"""
    try:
        body = json.dumps(response.json(), ensure_ascii=False, indent=2)
    except (ValueError, AttributeError):
        body = getattr(response, "text", "") or ""
    status_code = getattr(response, "status_code", None)
    if status_code is None:
        return body
    return f"HTTP {status_code}\n{body}"
