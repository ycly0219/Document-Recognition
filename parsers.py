"""三种单据模板的列头定义、识别结果解析与发票明细拆分。"""

import re

from logging_utils import print_log


# 订单类型选项集中维护，后续扩展只需增加对应模板的选项。
PICK_ORDER_TYPE_OPTIONS = (
    ("国内出库_FE", "GNCK_FE"),
    ("国内出库维修订单", "GNCK_WX"),
    ("国内出库报废订单", "GNCK_BF"),
    ("国内出库补货订单", "GNCK_BH"),
    ("国内出库大保养订单", "GNCK_DBY"),
    ("国外出库400", "GWCK_400"),
    ("国外出库600", "GWCK_600"),
    ("国外出库700", "GWCK_700"),
    ("国外出库900", "GWCK_900"),
)

INVOICE_ORDER_TYPE_OPTIONS = (
    ("国外入库", "OSI"),
    ("国内采购入库", "POIN"),
    ("国内外维修入库", "REPAIRIN"),
)

ORDER_TYPE_OPTIONS_BY_TEMPLATE = {
    "GE-ORACLE拣货单": PICK_ORDER_TYPE_OPTIONS,
    "GE-OSCAR拣货单": PICK_ORDER_TYPE_OPTIONS,
    "GE-发票单": INVOICE_ORDER_TYPE_OPTIONS,
}

DEFAULT_ORDER_TYPE_BY_TEMPLATE = {
    "GE-ORACLE拣货单": "",
    "GE-OSCAR拣货单": "",
    "GE-发票单": INVOICE_ORDER_TYPE_OPTIONS[0][0],
}


def get_order_type_labels(select_text):
    """返回指定模板可选订单类型中文标签。"""
    return [label for label, _ in ORDER_TYPE_OPTIONS_BY_TEMPLATE.get(select_text, ())]


def get_order_type_value(select_text, label):
    """把订单类型中文标签转换为当前模板导出值，未知值原样返回。"""
    mapping = dict(ORDER_TYPE_OPTIONS_BY_TEMPLATE.get(select_text, ()))
    return mapping.get(str(label).strip(), label)


def get_default_order_type_label(select_text):
    """返回指定模板初始订单类型；拣货单为空，需要人工必选。"""
    return DEFAULT_ORDER_TYPE_BY_TEMPLATE.get(select_text, "")


_ITEM_DETAIL_PATTERN = re.compile(
    r"(?<![A-Za-z])(LPN|Serial|Lot|COO):\s*(.*?)(?=\s*(?:LPN|Serial|Lot|COO):|\Z)",
    re.DOTALL,
)


def get_core_headers(select_text):
    """返回各单据模板对应的预览/导出列结构。"""
    if select_text == "GE-ORACLE拣货单":
        return [
            "订单类型",
            "Order Number",
            "Task Id",
            "Item Number",
            "Qty",
            "LPN",
            "Serial",
            "Lot",
            "COO",
            "Pick From Locator",
            "OrderType",
            "Ordered Date",
            "Shipment Priority",
            "Ship Method",
            "Service Level",
            "FE SSO",
            "FE Name",
            "Customer Name",
            "Customer Number",
            "SHIP TO NO",
            "Ship To Address",
            "Email",
            "Shipping Instruction",
            "Special Instruction",
            "Org",
            "Pick Slip Print Date",
            "System Id",
            "Pick From Subinv",
            "Customer PO",
            "Delivery",
        ]
    elif select_text == "GE-OSCAR拣货单":
        return [
            "订单类型",
            "服务申请号", "物料编号", "数量", "序列号", "货位",
            "状态", "仓库", "供应商", "SSO", "收货人",
            "姓名", "收货地址", "时效", "收货人电话", "申请说明",
            "客户设备id", "SR编号", "跟踪号"
        ]
    elif select_text == "GE-发票单":
        # 列头顺序同时用于预览表格与 Excel 导出
        return [
            "订单类型",
            "运单号",
            "INVOICE NO", "ITEM NUMBER", "QTY", "LPN Number", "Serial Number",
            "LOT Number", "Expiration Date", "COUNTRY OF ORIGIN",
            "SALES ORDER NO", "CUSTOMER PO", "DATE", "DELIVERY", "CARRIER", "HAWB"
        ]
    raise ValueError(f"未知模板: {select_text}")


_PREVIEW_LAYOUT = {
    "GE-ORACLE拣货单": (
        [
            "订单类型",
            "Order Number",
            "OrderType",
            "Pick Slip Print Date",
            "Shipment Priority",
            "Service Level",
            "FE SSO",
            "FE Name",
            "SHIP TO NO",
            "Ship To Address",
            "Shipping Instruction",
            "Special Instruction",
            "Customer Name",
            "Customer Number",
            "System Id",
            "Pick From Subinv",
            "Customer PO",
            "Delivery",
            "Ordered Date",
            "Ship Method",
        ],
        [
            "Task Id",
            "Item Number",
            "Qty",
            "LPN",
            "Serial",
            "Lot",
            "COO",
            "Pick From Locator",
            "Org",
        ],
    ),
    "GE-OSCAR拣货单": (
        [
            "订单类型",
            "服务申请号", "SR编号", "时效", "供应商", "收货人",
            "收货地址", "收货人电话", "申请说明", "SSO", "姓名", "客户设备id",
        ],
        [
            "物料编号", "数量", "序列号", "货位", "仓库", "状态", "跟踪号",
        ],
    ),
    "GE-发票单": (
        [
            "订单类型", "运单号", "INVOICE NO", "DATE", "DELIVERY", "CARRIER", "HAWB",
        ],
        [
            "ITEM NUMBER", "QTY", "LPN Number", "Serial Number",
            "LOT Number", "Expiration Date", "COUNTRY OF ORIGIN",
            "SALES ORDER NO", "CUSTOMER PO",
        ],
    ),
}


def get_preview_layout(select_text):
    """返回预览用的 Header 字段顺序与 Details 字段顺序。"""
    if select_text not in _PREVIEW_LAYOUT:
        raise ValueError(f"未知模板: {select_text}")
    return _PREVIEW_LAYOUT[select_text]


def merge_preview_rows(select_text, header_values, detail_rows):
    """按导出所需的完整列顺序，把单组单据头与明细行重组为完整行。"""
    full_headers = get_core_headers(select_text)
    header_fields, detail_fields = get_preview_layout(select_text)
    header_index = {name: index for index, name in enumerate(full_headers)}
    detail_index = {name: index for index, name in enumerate(detail_fields)}
    merged_rows = []
    for detail_row in detail_rows:
        row = [header_values.get(name, "") for name in full_headers]
        for name in detail_fields:
            index = detail_index[name]
            row[header_index[name]] = detail_row[index] if index < len(detail_row) else ""
        merged_rows.append(row)
    return merged_rows


def _parse_oracle_picklist(commit_result, filename):
    """解析 GE-ORACLE 拣货单，返回明细行列表。"""
    order_number = commit_result.get("Order Number", {}).get("value", "").strip()
    order_type = commit_result.get("OrderType", {}).get("value", "").strip()
    ordered_date = commit_result.get("Ordered Date", {}).get("value", "").strip()
    shipment_priority = commit_result.get("Shipment Priority", {}).get("value", "").strip()
    ship_method = commit_result.get("Ship Method", {}).get("value", "").strip()
    service_level = commit_result.get("Service Level", {}).get("value", "").strip()
    fe_sso = commit_result.get("FE SSO", {}).get("value", "").strip()
    fe_name = commit_result.get("FE Name", {}).get("value", "").strip()
    customer_name = commit_result.get("Customer Name", {}).get("value", "").strip()
    customer_number = commit_result.get("Customer Number", {}).get("value", "").strip()
    shipping_instruction = commit_result.get("Shipping Instruction", {}).get("value", "").strip()
    special_instruction = commit_result.get("Special Instruction", {}).get("value", "").strip()
    org = commit_result.get("Org", {}).get("value", "").strip()
    pick_slip_print_date = commit_result.get("Pick Slip Print Date", {}).get("value", "").strip()
    system_id = commit_result.get("System Id", {}).get("value", "").strip()
    pick_from_subinv = commit_result.get("Pick From Subinv", {}).get("value", "").strip()
    customer_po = commit_result.get("Customer PO", {}).get("value", "").strip()

    ship_addr = commit_result.get("Ship To Address", {}).get("value", "").strip()
    ship_to_no = commit_result.get("SHIP TO NO", {}).get("value", "").strip()
    email = commit_result.get("Email", {}).get("value", "").strip()
    delivery = commit_result.get("Delivery", {}).get("value", "").strip()
    material_list = commit_result.get("物料信息", {}).get("content", [])

    if not material_list:
        raise Exception("未识别到拣货明细数据")
    if not order_number:
        raise Exception("未识别到Order Number订单号")

    print_log(f"{filename} OrderNumber:{order_number} Delivery:{delivery} 明细行数:{len(material_list)}")
    rows = []
    for item in material_list:
        task_id = item.get("Task Id", {}).get("value", "").strip()
        item_no = item.get("Item Number", {}).get("value", "").strip()
        qty = item.get("Qty", {}).get("value", "").strip()
        pick_loc = item.get("Pick From Locator", {}).get("value", "").strip()
        item_detail = item.get("Item Details", {}).get("value", "").strip()
        un_number_index = item_detail.find("UN Number:")
        if un_number_index >= 0:
            item_detail = item_detail[:un_number_index]
        trace_values = {label: "" for label in ("LPN", "Serial", "Lot", "COO")}
        for label, value in _ITEM_DETAIL_PATTERN.findall(item_detail):
            trace_values[label] = value.strip()

        # 按 get_core_headers 的预览列顺序拼接字段
        rows.append([
            "",
            order_number,
            task_id,
            item_no,
            qty,
            trace_values["LPN"],
            trace_values["Serial"],
            trace_values["Lot"],
            trace_values["COO"],
            pick_loc,
            order_type,
            ordered_date,
            shipment_priority,
            ship_method,
            service_level,
            fe_sso,
            fe_name,
            customer_name,
            customer_number,
            ship_to_no,
            ship_addr,
            email,
            shipping_instruction,
            special_instruction,
            org,
            pick_slip_print_date,
            system_id,
            pick_from_subinv,
            customer_po,
            delivery,
        ])
    return rows, []


def _parse_oscar_picklist(commit_result, filename):
    """解析 GE-OSCAR 拣货单，返回明细行列表与空拆分分组。"""
    service_apply_no = commit_result.get("服务申请号", {}).get("value", "").strip()
    sr_no = commit_result.get("SR编号", {}).get("value", "").strip()
    supplier = commit_result.get("供应商", {}).get("value", "").strip()
    sso = commit_result.get("SSO", {}).get("value", "").strip()
    ship_addr = commit_result.get("收货地址", {}).get("value", "").strip()
    lead_time = commit_result.get("时效", {}).get("value", "").strip()
    consignee_tel = commit_result.get("收货人电话", {}).get("value", "").strip()
    apply_note = commit_result.get("申请说明", {}).get("value", "").strip()
    consignee_name = commit_result.get("收货人", {}).get("value", "").strip()
    real_name = commit_result.get("姓名", {}).get("value", "").strip()
    cust_device_id = commit_result.get("客户设备id", {}).get("value", "").strip()

    material_list = commit_result.get("物料信息", {}).get("content", [])

    if not material_list:
        raise Exception("未识别到OSCAR拣货明细数据")
    if not service_apply_no:
        raise Exception("未识别到服务申请号")

    print_log(f"{filename} 服务申请号:{service_apply_no} SR编号:{sr_no} 收货人:{consignee_name} 姓名:{real_name} 客户设备id:{cust_device_id} 明细行数:{len(material_list)}")
    rows = []
    for item in material_list:
        mat_no = item.get("物料编号", {}).get("value", "").strip()
        qty = item.get("数量", {}).get("value", "").strip()
        serial_no = item.get("序列号", {}).get("value", "").strip()
        track_no = item.get("跟踪号", {}).get("value", "").strip()
        locator = item.get("货位", {}).get("value", "").strip()
        status_val = item.get("状态", {}).get("value", "").strip()
        warehouse = item.get("仓库", {}).get("value", "").strip()

        rows.append([
            "",
            service_apply_no, mat_no, qty, serial_no, locator, status_val,
            warehouse, supplier, sso, consignee_name, real_name, ship_addr,
            lead_time, consignee_tel, apply_note, cust_device_id, sr_no,
            track_no,
        ])
    return rows, []


def _parse_invoice(commit_result, filename):
    """解析 GE-发票单，并按 LPN/Serial 与数量关系拆分明细。

    返回 (明细行列表, 拆分汇总元数据)；明细行仍保持可导出的平铺结构；
    只拆出 1 条子行时不返回汇总元数据。
    """
    order_type = get_default_order_type_label("GE-发票单")
    tracking_no = ""
    invoice_no = commit_result.get("INVOICE NO", {}).get("value", "").strip()
    delivery = commit_result.get("DELIVERY", {}).get("value", "").strip()
    doc_date = commit_result.get("DATE", {}).get("value", "").strip()
    carrier = commit_result.get("CARRIER", {}).get("value", "").strip()
    hawb = commit_result.get("HAWB", {}).get("value", "").strip()

    material_list = commit_result.get("物料信息", {}).get("content", [])

    if not material_list:
        raise Exception("未识别到发票明细数据")
    if not invoice_no:
        raise Exception("未识别到INVOICE NO发票号")

    print_log(f"{filename} INVOICE NO:{invoice_no} DELIVERY:{delivery} DATE:{doc_date} CARRIER:{carrier} HAWB:{hawb} 明细总行数:{len(material_list)}")
    rows = []
    split_groups = []
    for source_index, item in enumerate(material_list):
        raw_qty_str = item.get("QTY", {}).get("value", "").strip()
        item_num = item.get("ITEM NUMBER", {}).get("value", "").strip()
        country_of_origin = item.get("COUNTRY OF ORIGIN", {}).get("value", "").strip()
        raw_lpn_str = item.get("LPN Number", {}).get("value", "").strip()
        raw_serial_str = item.get("Serial Number", {}).get("value", "").strip()
        lot = item.get("LOT Number", {}).get("value", "").strip()
        sales_order = item.get("SALES ORDER NO", {}).get("value", "").strip()
        customer_po = item.get("CUSTOMER PO", {}).get("value", "").strip()
        expire = item.get("Expiration Date", {}).get("value", "").strip()
        split_summary_row = [
            order_type, tracking_no, invoice_no, item_num, raw_qty_str,
            "原始行汇总", "", lot, expire,
            country_of_origin, sales_order, customer_po,
            doc_date, delivery, carrier, hawb
        ]

        # 清洗LPN列表
        lpn_list = [lpn.strip() for lpn in raw_lpn_str.split(",") if lpn.strip()]
        lpn_count = len(lpn_list)
        # 清洗Serial列表（为空则返回空数组）
        serial_list = [s.strip() for s in raw_serial_str.split(",") if s.strip()]
        serial_count = len(serial_list)
        # 转换QTY数字，异常置0
        try:
            qty_val = int(raw_qty_str)
        except ValueError:
            qty_val = 0

        # ---------------- 拆分规则（兼容Serial为空） ----------------
        # 情况1：Serial为空，仅按LPN原有规则拆分，Serial填空
        if serial_count == 0:
            if lpn_count == qty_val and lpn_count > 0:
                print_log(f"Serial为空，匹配LPN规则1：LPN数量={lpn_count}=QTY{qty_val}，逐个拆分")
                child_indexes = []
                for single_lpn in lpn_list:
                    child_indexes.append(len(rows))
                    rows.append([
                        order_type, tracking_no, invoice_no, item_num, "1", single_lpn, "", lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
                # 单条拆分直接作为普通明细行，不生成汇总分组
                if len(child_indexes) > 1:
                    split_groups.append({
                        "source_index": source_index,
                        "source_qty": raw_qty_str,
                        "summary_row": split_summary_row,
                        "child_indexes": child_indexes,
                    })
            elif lpn_count == 1 and qty_val > 0:
                single_lpn = lpn_list[0]
                print_log(f"Serial为空，匹配LPN规则2：单LPN，QTY={qty_val}，保留一行")
                rows.append([
                    order_type, tracking_no, invoice_no, item_num, raw_qty_str, single_lpn, "", lot, expire,
                    country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
            else:
                # LPN不满足拆分，原样一行
                rows.append([
                    order_type, tracking_no, invoice_no, item_num, raw_qty_str, raw_lpn_str, "", lot, expire,
                    country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
        # 情况2：Serial有值，原有多字段匹配逻辑
        else:
            # 场景1：LPN数量=QTY，Serial数量也等于QTY，一一对应拆分
            if lpn_count == qty_val and serial_count == qty_val and qty_val > 0:
                print_log(f"匹配规则1：LPN/Serial数量={lpn_count}=QTY{qty_val}，逐行匹配拆分")
                child_indexes = []
                for idx in range(qty_val):
                    single_lpn = lpn_list[idx]
                    single_serial = serial_list[idx]
                    child_indexes.append(len(rows))
                    rows.append([
                        order_type, tracking_no, invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
                # 单条拆分直接作为普通明细行，不生成汇总分组
                if len(child_indexes) > 1:
                    split_groups.append({
                        "source_index": source_index,
                        "source_qty": raw_qty_str,
                        "summary_row": split_summary_row,
                        "child_indexes": child_indexes,
                    })
            # 场景2：仅1个LPN、仅1个Serial，按QTY循环复制N行
            elif lpn_count == 1 and serial_count == 1 and qty_val > 0:
                single_lpn = lpn_list[0]
                single_serial = serial_list[0]
                print_log(f"匹配规则2：单LPN+单Serial，QTY={qty_val}，保留一行")
                rows.append([
                    order_type, tracking_no, invoice_no, item_num, raw_qty_str, single_lpn, single_serial, lot, expire,
                    country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
            # 场景3：只有Serial多个、LPN单个，且serial_count == qty_val
            elif lpn_count == 1 and serial_count == qty_val and qty_val > 0:
                single_lpn = lpn_list[0]
                print_log(f"匹配规则3：单LPN，Serial数量={serial_count}=QTY{qty_val}，拆分Serial多行")
                child_indexes = []
                for single_serial in serial_list:
                    child_indexes.append(len(rows))
                    rows.append([
                        order_type, tracking_no, invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
                split_groups.append({
                    "source_index": source_index,
                    "source_qty": raw_qty_str,
                    "summary_row": split_summary_row,
                    "child_indexes": child_indexes,
                })
            # 场景4：只有LPN多个、Serial单个，且lpn_count == qty_val
            elif serial_count == 1 and lpn_count == qty_val and qty_val > 0:
                single_serial = serial_list[0]
                print_log(f"匹配规则4：单Serial，LPN数量={lpn_count}=QTY{qty_val}，拆分LPN多行")
                child_indexes = []
                for single_lpn in lpn_list:
                    child_indexes.append(len(rows))
                    rows.append([
                        order_type, tracking_no, invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
                split_groups.append({
                    "source_index": source_index,
                    "source_qty": raw_qty_str,
                    "summary_row": split_summary_row,
                    "child_indexes": child_indexes,
                })
            # 其他所有不匹配场景，保留原始一行，LPN/Serial逗号拼接不拆分
            else:
                rows.append([
                    order_type, tracking_no, invoice_no, item_num, raw_qty_str, raw_lpn_str, raw_serial_str,
                    lot, expire, country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
    return rows, split_groups


def parse_commit_result(select_text, commit_result, filename):
    """按当前模板解析单份 OCR 识别结果，返回 (明细行列表, 拆分汇总元数据)。"""
    if select_text == "GE-ORACLE拣货单":
        return _parse_oracle_picklist(commit_result, filename)
    elif select_text == "GE-OSCAR拣货单":
        return _parse_oscar_picklist(commit_result, filename)
    elif select_text == "GE-发票单":
        return _parse_invoice(commit_result, filename)
    raise ValueError(f"未知模板: {select_text}")
