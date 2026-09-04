"""Flux WMS putPurchaseOrder / putOriginalSalesOrder 报文构建与发送。"""

from datetime import datetime
import json

import requests

from config import (
    WMS_CUSTOMER_ID,
    WMS_PUT_ORIGINAL_SALES_ORDER_URL,
    WMS_PUT_PURCHASE_ORDER_URL,
    WMS_WAREHOUSE_ID,
)
from excel_export import (
    _normalize_expiration_date,
    _normalize_oracle_datetime,
    _normalize_oscar_serial,
    _oracle_quality_status,
)
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

_ORACLE_DETAIL_HEADERS = [
    "Task Id",
    "Item Number",
    "Qty",
    "LPN",
    "Serial",
    "Lot",
    "COO",
    "Pick From Locator",
    "Org",
]

_OSCAR_DETAIL_HEADERS = [
    "物料编号",
    "数量",
    "序列号",
    "货位",
    "仓库",
    "状态",
    "跟踪号",
]

CONSIGNEE_NAME = "虚拟收货人"


def _text(value):
    """把界面值统一转成 WMS 报文使用的字符串。"""
    return "" if value is None else str(value).strip()


def _optional_item(target, key, value):
    """仅在有值时写入可选报文字段，空值不进入 JSON。"""
    text = _text(value)
    if text:
        target[key] = text


def _build_oracle_put_original_sales_order_payload(header_values, detail_rows):
    """按 GE-ORACLE 拣货单预览字段构建 putOriginalSalesOrder 报文。"""
    header = {
        "warehouseId": WMS_WAREHOUSE_ID,
        "customerId": WMS_CUSTOMER_ID,
        "consigneeId": "CONSIGNEEID",
        "consigneeName": CONSIGNEE_NAME,
    }
    _optional_item(
        header,
        "orderType",
        get_order_type_value(
            "GE-ORACLE拣货单", header_values.get("订单类型", "")
        ),
    )
    _optional_item(header, "docNo", header_values.get("Order Number"))
    _optional_item(header, "soReferenceB", header_values.get("System Id"))
    _optional_item(
        header,
        "orderTime",
        _normalize_oracle_datetime(
            header_values.get("Pick Slip Print Date"), include_time=True
        ),
    )
    _optional_item(
        header, "consigneeAddress1", header_values.get("Ship To Address")
    )
    _optional_item(header, "hedi01", header_values.get("OrderType"))
    _optional_item(
        header,
        "hedi02",
        _normalize_oracle_datetime(
            header_values.get("Ordered Date"), include_time=False
        ),
    )
    _optional_item(
        header, "hedi03", header_values.get("Shipment Priority")
    )
    _optional_item(header, "hedi04", header_values.get("Ship Method"))
    _optional_item(header, "hedi05", header_values.get("Service Level"))
    _optional_item(header, "hedi06", header_values.get("FE SSO"))
    _optional_item(header, "hedi07", header_values.get("FE Name"))
    _optional_item(
        header, "hedi08", header_values.get("Shipping Instruction")
    )
    _optional_item(
        header, "hedi11", header_values.get("Special Instruction")
    )
    _optional_item(
        header, "hedi12", header_values.get("Pick From Subinv")
    )
    _optional_item(
        header, "hedi13", header_values.get("Pick From Subinv")
    )
    _optional_item(header, "userDefine1", header_values.get("SHIP TO NO"))

    pick_from_subinv = header_values.get("Pick From Subinv")
    details = []
    for line_no, row in enumerate(detail_rows or [], start=1):
        row_map = dict(zip(_ORACLE_DETAIL_HEADERS, row))
        detail = {
            "lineNo": str(line_no),
            "sku": _text(row_map.get("Item Number")),
            "qtyOrdered": _text(row_map.get("Qty")),
            "packUom": "EA",
            "price": "0",
            "lotAtt05": "ORACLE",
        }
        _optional_item(detail, "lotAtt04", row_map.get("Lot"))
        _optional_item(detail, "lotAtt07", row_map.get("Org"))
        _optional_item(
            detail, "lotAtt08", _oracle_quality_status(pick_from_subinv)
        )
        _optional_item(detail, "lotAtt09", row_map.get("Serial"))
        _optional_item(detail, "lotAtt11", row_map.get("LPN"))
        _optional_item(detail, "dedi01", row_map.get("Task Id"))
        _optional_item(detail, "dedi03", row_map.get("Pick From Locator"))
        details.append(detail)

    return {"data": {"header": [{**header, "details": details}]}}


def _build_oscar_put_original_sales_order_payload(header_values, detail_rows):
    """按 GE-OSCAR 拣货单预览字段构建 putOriginalSalesOrder 报文。"""
    header = {
        "warehouseId": WMS_WAREHOUSE_ID,
        "customerId": WMS_CUSTOMER_ID,
        "consigneeId": "CONSIGNEEID",
        "consigneeName": CONSIGNEE_NAME,
    }
    _optional_item(
        header,
        "orderType",
        get_order_type_value(
            "GE-OSCAR拣货单", header_values.get("订单类型", "")
        ),
    )
    _optional_item(header, "docNo", header_values.get("服务申请号"))
    _optional_item(header, "soReferenceA", header_values.get("SR编号"))
    _optional_item(header, "soReferenceB", header_values.get("客户设备id"))
    header["orderTime"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _optional_item(header, "notes", header_values.get("申请说明"))
    _optional_item(
        header, "consigneeAddress1", header_values.get("收货地址")
    )
    _optional_item(header, "hedi01", header_values.get("客户设备id"))
    _optional_item(header, "hedi05", header_values.get("时效"))
    _optional_item(header, "hedi06", header_values.get("SSO"))
    _optional_item(header, "hedi07", header_values.get("姓名"))
    _optional_item(header, "hedi13", header_values.get("供应商"))
    _optional_item(header, "hedi14", header_values.get("收货人"))
    _optional_item(header, "hedi15", header_values.get("收货人电话"))

    details = []
    for line_no, row in enumerate(detail_rows or [], start=1):
        row_map = dict(zip(_OSCAR_DETAIL_HEADERS, row))
        detail = {
            "lineNo": str(line_no),
            "sku": _text(row_map.get("物料编号")),
            "qtyOrdered": _text(row_map.get("数量")),
            "packUom": "EA",
            "price": "0",
            "lotAtt05": "OSCAR",
        }
        _optional_item(detail, "lotAtt07", row_map.get("仓库"))
        if _text(row_map.get("状态")) == "好件":
            detail["lotAtt08"] = "GOOD"
        _optional_item(
            detail,
            "lotAtt09",
            _normalize_oscar_serial(row_map.get("序列号")),
        )
        _optional_item(detail, "dedi02", row_map.get("跟踪号"))
        _optional_item(detail, "dedi03", row_map.get("货位"))
        details.append(detail)

    return {"data": {"header": [{**header, "details": details}]}}


def build_put_original_sales_order_payload(select_text, header_values, detail_rows):
    """按拣货单模板构建 putOriginalSalesOrder 报文。"""
    if select_text == "GE-ORACLE拣货单":
        return _build_oracle_put_original_sales_order_payload(
            header_values, detail_rows
        )
    if select_text == "GE-OSCAR拣货单":
        return _build_oscar_put_original_sales_order_payload(
            header_values, detail_rows
        )
    raise ValueError(f"不支持的销售订单模板: {select_text}")


def build_put_purchase_order_payload(header_values, detail_rows):
    """按发票页签的单据头和可导出发明细构建 putPurchaseOrder 报文。"""
    header = {
        "warehouseId": WMS_WAREHOUSE_ID,
        "customerId": WMS_CUSTOMER_ID,
        "poType": get_order_type_value(
            "GE-发票单", header_values.get("订单类型", "")
        ),
        "docNo": _text(header_values.get("INVOICE NO", "")),
        "poReferenceA": _text(header_values.get("运单号", "")),
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


def send_put_original_sales_order(payload):
    """发送 putOriginalSalesOrder 报文并返回接口响应对象。"""
    return requests.post(
        WMS_PUT_ORIGINAL_SALES_ORDER_URL, json=payload, timeout=30
    )


def is_wms_send_success(response):
    """按 HTTP 状态与顶层 returnFlag 判定 WMS 接口是否发送成功。"""
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
