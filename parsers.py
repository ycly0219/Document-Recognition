"""三种单据模板的列头定义、识别结果解析与发票明细拆分。"""

from logging_utils import print_log


def get_core_headers(select_text):
    """返回各单据模板对应的预览/导出列结构。"""
    if select_text == "GE-ORACLE拣货单":
        # Order Number后面紧跟全部新增字段
        return [
            "Order Number",
            "OrderType",
            "Ordered Date",
            "Shipment Priority",
            "Ship Method",
            "Service Level",
            "FE SSO",
            "FE Name",
            "Customer Name",
            "Customer Number",
            "Shipping Instruction",
            "Special Instruction",
            "Org",
            "Pick Slip Print Date",
            "System Id",
            "Pick From Subinv",
            "Customer PO",
            "Ship To Address",
            "Email",
            "Delivery",
            "Task Id",
            "Item Number",
            "Qty",
            "Pick From Locator",
            "Item Details"
        ]
    elif select_text == "GE-OSCAR拣货单":
        return [
            "服务申请号", "SR编号", "供应商", "SSO", "收货地址",
            "时效", "收货人电话", "申请说明",
            "收货人", "姓名", "客户设备id",
            "物料编号", "数量", "序列号", "跟踪号", "货位", "状态", "仓库"
        ]
    elif select_text == "GE-发票单":
        # 列头顺序同时用于预览表格与 Excel 导出
        return [
            "INVOICE NO", "ITEM NUMBER", "QTY", "LPN Number", "Serial Number",
            "LOT Number", "Expiration Date", "COUNTRY OF ORIGIN",
            "SALES ORDER NO", "CUSTOMER PO", "DATE", "DELIVERY", "CARRIER", "HAWB"
        ]
    raise ValueError(f"未知模板: {select_text}")


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

        # Order Number后面拼接全部新增字段
        rows.append([
            order_number,
            order_type,
            ordered_date,
            shipment_priority,
            ship_method,
            service_level,
            fe_sso,
            fe_name,
            customer_name,
            customer_number,
            shipping_instruction,
            special_instruction,
            org,
            pick_slip_print_date,
            system_id,
            pick_from_subinv,
            customer_po,
            ship_addr,
            email,
            delivery,
            task_id,
            item_no,
            qty,
            pick_loc,
            item_detail
        ])
    return rows


def _parse_oscar_picklist(commit_result, filename):
    """解析 GE-OSCAR 拣货单，返回明细行列表。"""
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
            service_apply_no, sr_no, supplier, sso, ship_addr,
            lead_time, consignee_tel, apply_note,
            consignee_name, real_name, cust_device_id,
            mat_no, qty, serial_no, track_no, locator, status_val, warehouse
        ])
    return rows


def _parse_invoice(commit_result, filename):
    """解析 GE-发票单，并按 LPN/Serial 与数量关系拆分明细。"""
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
    for item in material_list:
        raw_qty_str = item.get("QTY", {}).get("value", "").strip()
        item_num = item.get("ITEM NUMBER", {}).get("value", "").strip()
        country_of_origin = item.get("COUNTRY OF ORIGIN", {}).get("value", "").strip()
        raw_lpn_str = item.get("LPN Number", {}).get("value", "").strip()
        raw_serial_str = item.get("Serial Number", {}).get("value", "").strip()
        lot = item.get("LOT Number", {}).get("value", "").strip()
        sales_order = item.get("SALES ORDER NO", {}).get("value", "").strip()
        customer_po = item.get("CUSTOMER PO", {}).get("value", "").strip()
        expire = item.get("Expiration Date", {}).get("value", "").strip()

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
                for single_lpn in lpn_list:
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, "", lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            elif lpn_count == 1 and qty_val > 0:
                split_row_count = qty_val
                single_lpn = lpn_list[0]
                print_log(f"Serial为空，匹配LPN规则2：单LPN，QTY={qty_val}，生成{split_row_count}条")
                for _ in range(split_row_count):
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, "", lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            else:
                # LPN不满足拆分，原样一行
                rows.append([
                    invoice_no, item_num, raw_qty_str, raw_lpn_str, "", lot, expire,
                    country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
        # 情况2：Serial有值，原有多字段匹配逻辑
        else:
            # 场景1：LPN数量=QTY，Serial数量也等于QTY，一一对应拆分
            if lpn_count == qty_val and serial_count == qty_val and qty_val > 0:
                print_log(f"匹配规则1：LPN/Serial数量={lpn_count}=QTY{qty_val}，逐行匹配拆分")
                for idx in range(qty_val):
                    single_lpn = lpn_list[idx]
                    single_serial = serial_list[idx]
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            # 场景2：仅1个LPN、仅1个Serial，按QTY循环复制N行
            elif lpn_count == 1 and serial_count == 1 and qty_val > 0:
                split_row_count = qty_val
                single_lpn = lpn_list[0]
                single_serial = serial_list[0]
                print_log(f"匹配规则2：单LPN+单Serial，QTY={qty_val}，生成{split_row_count}条重复数据")
                for _ in range(split_row_count):
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            # 场景3：只有Serial多个、LPN单个，且serial_count == qty_val
            elif lpn_count == 1 and serial_count == qty_val and qty_val > 0:
                single_lpn = lpn_list[0]
                print_log(f"匹配规则3：单LPN，Serial数量={serial_count}=QTY{qty_val}，拆分Serial多行")
                for single_serial in serial_list:
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            # 场景4：只有LPN多个、Serial单个，且lpn_count == qty_val
            elif serial_count == 1 and lpn_count == qty_val and qty_val > 0:
                single_serial = serial_list[0]
                print_log(f"匹配规则4：单Serial，LPN数量={lpn_count}=QTY{qty_val}，拆分LPN多行")
                for single_lpn in lpn_list:
                    rows.append([
                        invoice_no, item_num, "1", single_lpn, single_serial, lot, expire,
                        country_of_origin, sales_order, customer_po,
                        doc_date, delivery, carrier, hawb
                    ])
            # 其他所有不匹配场景，保留原始一行，LPN/Serial逗号拼接不拆分
            else:
                rows.append([
                    invoice_no, item_num, raw_qty_str, raw_lpn_str, raw_serial_str,
                    lot, expire, country_of_origin, sales_order, customer_po,
                    doc_date, delivery, carrier, hawb
                ])
    return rows


def parse_commit_result(select_text, commit_result, filename):
    """按当前模板解析单份 OCR 识别结果，返回明细行列表。"""
    if select_text == "GE-ORACLE拣货单":
        return _parse_oracle_picklist(commit_result, filename)
    elif select_text == "GE-OSCAR拣货单":
        return _parse_oscar_picklist(commit_result, filename)
    elif select_text == "GE-发票单":
        return _parse_invoice(commit_result, filename)
    raise ValueError(f"未知模板: {select_text}")
