import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import requests
import json
import logging
import os
import queue
import threading
import time
import sys
from openpyxl import Workbook

# ================== 配置区域 ==================
# OCR 相关配置
API_URL = "http://172.30.197.3:10000/OcrPlugins/generalOcr/extractGeneralDataAsync"
UPLOAD_API_URL = "http://techfile.i.sinotrans.com:80/objectstorecloud/files/v2"
GET_RESULT_API_URL = "http://172.30.197.3:10000/OcrPlugins/customsOcr/getResultByReqUuid"

ORG_ID = "99999"
SOURCE_CODE = "99999"

OCR_APP_ID = "R5QHOMZd"
OCR_APP_SECRET = "c108525f68088809f53b3ed715abd826"
OCR_APP_KEY = "1MsDKlaQ"
OCR_SYS_CODE = "99999"   # 默认sysCode，非OSCAR单据使用
OCR_SYS_CODE_OSCAR = "LOGISTICS_GE"  # OSCAR拣货单专用sysCode
OCR_ORG_ID = "101162"
OCR_REGIONAL_CODE = ""
OCR_DOC_TYPE = "SINGLE_LLM_EXTRACT_ASYNC"
OCR_CALLBACK_URL = "http://172.30.254.38:10000/OcrPlugins/customsOcr/test"
OCR_MAX_RETRY = 50
OCR_RETRY_INTERVAL = 3

# 模板映射（GE拣货单、GE‑OSCAR拣货单、GE‑发票单）
MODEL_MAP = {
    "GE-ORACLE拣货单": "logistics_east_ge_picklist_99999_1503",
    "GE-OSCAR拣货单": "logistics_ge_oscarpicklist_1503",
    "GE-发票单": "logistics_east_ge_invoice_99999_1503"
}

# ---------------- 飞书多维表格配置 ----------------
FEISHU_APP_ID = "cli_aa978beae8f81cca"
FEISHU_APP_SECRET = "ywHBY0AmJc00TojMIghLzgRHpyngHXpR"
BITABLE_RECORDS_URL = "https://open.feishu.cn/open-apis/bitable/v1/apps/Jqfwbt2bWaLq7AswNz2c1iSUn4i/tables/tblQRssVXgCA7mEs/records"
# =============================================================

# 日志统一走标准 logging；GUI 日志通过队列由主线程刷新
log_queue = queue.Queue()
ui_message_queue = queue.Queue()
worker_thread = None

class GuiLogHandler(logging.Handler):
    def emit(self, record):
        log_queue.put(self.format(record))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
gui_handler = GuiLogHandler()
gui_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
logging.getLogger().addHandler(gui_handler)

def print_log(msg):
    """统一日志入口，写入标准 logging 和 GUI 日志队列。"""
    logging.info(msg)

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

def get_file_type_and_content_type(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".png":
        return "png", "image/png"
    elif ext in (".jpg", ".jpeg"):
        return "jpg", "image/jpeg"
    elif ext == ".pdf":
        return "pdf", "application/pdf"
    else:
        return ext.lstrip("."), "application/octet-stream"


def upload_file_to_server(file_path):
    filename = os.path.basename(file_path)
    print_log(f"正在上传文件: {filename}")
    try:
        file_type, content_type = get_file_type_and_content_type(file_path)
        with open(file_path, 'rb') as f:
            file_data = f.read()
        print_log(f"文件读取完成，大小 {len(file_data)//1024}KB，格式 {file_type}")

        files = {'file': (filename, file_data, content_type)}
        data = {'org_id': ORG_ID, 'source_code': SOURCE_CODE, 'file_type': file_type}

        response = requests.post(UPLOAD_API_URL, files=files, data=data, timeout=60)
        print_log(f"上传接口状态码: {response.status_code}")
        response.raise_for_status()
        result = response.json()

        if result.get('status') is True:
            file_id = result.get('fileId')
            file_path_ret = result.get('filePath')
            print_log(f"上传成功 fileId={file_id}")
            return file_id, file_path_ret
        else:
            raise Exception(f"接口返回失败: {result.get('message', '未知错误')}")
    except Exception as e:
        print_log(f"文件上传失败: {str(e)}")
        raise Exception(f"文件上传失败: {str(e)}")


def call_process_api(file_url, filename, file_id, model_id, rule_name):
    # 根据单据类型选择sysCode
    if rule_name == "GE-OSCAR拣货单":
        use_sys_code = OCR_SYS_CODE_OSCAR
        print_log(f"当前为GE‑OSCAR拣货单，使用sysCode={use_sys_code}")
    else:
        use_sys_code = OCR_SYS_CODE
        print_log(f"当前单据，使用sysCode={use_sys_code}")

    print_log(f"选用模型ID: {model_id}，提交OCR任务")
    payload = {
        "appId": OCR_APP_ID, "appSecret": OCR_APP_SECRET, "appKey": OCR_APP_KEY,
        "sysCode": use_sys_code, "orgId": OCR_ORG_ID, "regionalCode": OCR_REGIONAL_CODE,
        "docType": OCR_DOC_TYPE, "callBackUrl": OCR_CALLBACK_URL, "modelId": model_id,
        "files": [{"fileName": filename, "fileId": file_id}]
    }
    try:
        response = requests.post(API_URL, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, timeout=30)
        print_log(f"OCR提交接口状态码: {response.status_code}")
        response.raise_for_status()
        res = response.json()
        req_uuid = res.get("data", {}).get("reqUuid", "")
        if not req_uuid:
            raise Exception("OCR接口未返回reqUuid")
        print_log(f"获取任务reqUuid={req_uuid}")
        return req_uuid, ""
    except Exception as e:
        print_log(f"OCR提交接口失败: {str(e)}")
        raise Exception(f"OCR接口调用失败: {str(e)}")


def call_get_result_api(req_uuid):
    max_retry = OCR_MAX_RETRY
    retry_interval = OCR_RETRY_INTERVAL
    current_retry = 0

    while current_retry < max_retry:
        current_retry += 1
        print_log(f"第{current_retry}/{max_retry}次查询OCR结果 reqUuid={req_uuid}")
        try:
            params = {"reqUuid": req_uuid}
            response = requests.get(GET_RESULT_API_URL, params=params, timeout=30)
            response.raise_for_status()
            res_dict = response.json()

            if not res_dict.get("status"):
                print_log("接口status=false，等待重试")
                time.sleep(retry_interval)
                continue

            commit_result = res_dict.get("data", {}).get("commitResult")
            if commit_result is not None and commit_result != {}:
                print_log("OCR识别结果获取成功")
                return "", res_dict
            else:
                print_log("识别结果未生成，等待3秒")
                time.sleep(retry_interval)

        except Exception as e:
            print_log(f"第{current_retry}次查询异常: {str(e)}")
            time.sleep(retry_interval)

    raise Exception(f"查询{max_retry}次无有效识别结果")


# ---------------- 飞书相关函数 ----------------
def get_tenant_access_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    data = {
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        res = response.json()
        if res.get("code") == 0:
            print_log("获取飞书Token成功")
            return res["tenant_access_token"]
        else:
            print_log(f"获取飞书Token失败: {res['msg']}")
            return None
    except Exception as e:
        print_log(f"获取飞书Token异常: {str(e)}")
        return None


def send_to_bitable(token, rule_name, total_count):
    if not token:
        print_log("无有效Token，跳过写入飞书多维表格")
        return False

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    data = {
        "fields": {
            "单据模版规则": rule_name,
            "数量": total_count
        }
    }
    try:
        response = requests.post(BITABLE_RECORDS_URL, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        print_log("统计数据写入飞书多维表格完成")
        return True
    except Exception as e:
        print_log(f"写入飞书表格失败: {str(e)}")
        return False


# ---------------- 主流程函数 ----------------
def run_task():
    select_text = combo_model.get().strip()
    if not select_text or select_text not in MODEL_MAP:
        messagebox.showwarning("温馨提示", "请先选择单据规则！")
        return

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

    btn.config(text="正在处理...", state=tk.DISABLED)
    win.update()
    current_model_id = MODEL_MAP[select_text]

    global worker_thread
    worker_thread = threading.Thread(
        target=process_batch_worker,
        args=(files, select_text, current_model_id),
        daemon=True
    )
    worker_thread.start()
    win.after(100, poll_ui_queue)


def process_batch_worker(files, select_text, current_model_id):
    try:
        process_batch(files, select_text, current_model_id)
    except Exception as e:
        print_log(f"处理流程异常: {str(e)}")
        ui_message_queue.put(("error", f"处理流程异常: {str(e)}"))


def process_batch(files, select_text, current_model_id):
    total_files = len(files)
    full_log_data = []
    core_data = []
    success_count = 0
    fail_count = 0
    print_log(f"===== 批量处理，共{total_files}个文件，模版：{select_text} =====")

    for file_path in files:
        filename = os.path.basename(file_path)
        status = "失败"
        res_msg = ""
        file_id = ""
        file_url = ""
        req_uuid = ""
        ocr_result_dict = None

        try:
            file_id, file_url = upload_file_to_server(file_path)
            # 传入当前选中的单据规则名称，用于内部判断sysCode
            req_uuid, _ = call_process_api(file_url, filename, file_id, current_model_id, select_text)
            _, ocr_result_dict = call_get_result_api(req_uuid)

            if ocr_result_dict and ocr_result_dict.get("status") is True:
                commit_result = ocr_result_dict.get("data", {}).get("commitResult", {})

                # ================== GE‑ORACLE拣货单 解析逻辑【已扩展字段】 ==================
                if select_text == "GE-ORACLE拣货单":
                    order_number = commit_result.get("Order Number", {}).get("value", "").strip()
                    # 新增头部字段
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
                    for item in material_list:
                        task_id = item.get("Task Id", {}).get("value", "").strip()
                        item_no = item.get("Item Number", {}).get("value", "").strip()
                        qty = item.get("Qty", {}).get("value", "").strip()
                        pick_loc = item.get("Pick From Locator", {}).get("value", "").strip()
                        item_detail = item.get("Item Details", {}).get("value", "").strip()

                        # Order Number后面拼接全部新增字段
                        core_data.append([
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

                # ================== GE‑OSCAR拣货单【已新增：收货人、姓名、客户设备id；明细仓库】 ==================
                elif select_text == "GE-OSCAR拣货单":
                    service_apply_no = commit_result.get("服务申请号", {}).get("value", "").strip()
                    sr_no = commit_result.get("SR编号", {}).get("value", "").strip()
                    supplier = commit_result.get("供应商", {}).get("value", "").strip()
                    sso = commit_result.get("SSO", {}).get("value", "").strip()
                    ship_addr = commit_result.get("收货地址", {}).get("value", "").strip()
                    lead_time = commit_result.get("时效", {}).get("value", "").strip()
                    consignee_tel = commit_result.get("收货人电话", {}).get("value", "").strip()
                    apply_note = commit_result.get("申请说明", {}).get("value", "").strip()
                    # 新增头部字段
                    consignee_name = commit_result.get("收货人", {}).get("value", "").strip()
                    real_name = commit_result.get("姓名", {}).get("value", "").strip()
                    cust_device_id = commit_result.get("客户设备id", {}).get("value", "").strip()

                    material_list = commit_result.get("物料信息", {}).get("content", [])

                    if not material_list:
                        raise Exception("未识别到OSCAR拣货明细数据")
                    if not service_apply_no:
                        raise Exception("未识别到服务申请号")

                    print_log(f"{filename} 服务申请号:{service_apply_no} SR编号:{sr_no} 收货人:{consignee_name} 姓名:{real_name} 客户设备id:{cust_device_id} 明细行数:{len(material_list)}")
                    for item in material_list:
                        mat_no = item.get("物料编号", {}).get("value", "").strip()
                        qty = item.get("数量", {}).get("value", "").strip()
                        serial_no = item.get("序列号", {}).get("value", "").strip()
                        track_no = item.get("跟踪号", {}).get("value", "").strip()
                        locator = item.get("货位", {}).get("value", "").strip()
                        status_val = item.get("状态", {}).get("value", "").strip()
                        warehouse = item.get("仓库", {}).get("value", "").strip()

                        core_data.append([
                            service_apply_no, sr_no, supplier, sso, ship_addr,
                            lead_time, consignee_tel, apply_note,
                            consignee_name, real_name, cust_device_id,
                            mat_no, qty, serial_no, track_no, locator, status_val, warehouse
                        ])

                # ================== GE‑发票单 【增加DATE、CARRIER、HAWB】 ==================
                elif select_text == "GE-发票单":
                    invoice_no = commit_result.get("INVOICE NO", {}).get("value", "").strip()
                    delivery = commit_result.get("DELIVERY", {}).get("value", "").strip()
                    # 新增三个单据头字段
                    doc_date = commit_result.get("DATE", {}).get("value", "").strip()
                    carrier = commit_result.get("CARRIER", {}).get("value", "").strip()
                    hawb = commit_result.get("HAWB", {}).get("value", "").strip()

                    material_list = commit_result.get("物料信息", {}).get("content", [])

                    if not material_list:
                        raise Exception("未识别到发票明细数据")
                    if not invoice_no:
                        raise Exception("未识别到INVOICE NO发票号")

                    print_log(f"{filename} INVOICE NO:{invoice_no} DELIVERY:{delivery} DATE:{doc_date} CARRIER:{carrier} HAWB:{hawb} 明细总行数:{len(material_list)}")
                    for item in material_list:
                        raw_qty_str = item.get("QTY", {}).get("value", "").strip()
                        item_num = item.get("ITEM NUMBER", {}).get("value", "").strip()
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
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, "", lot, sales_order, customer_po, expire
                                    ])
                            elif lpn_count == 1 and qty_val > 0:
                                split_row_count = qty_val
                                single_lpn = lpn_list[0]
                                print_log(f"Serial为空，匹配LPN规则2：单LPN，QTY={qty_val}，生成{split_row_count}条")
                                for _ in range(split_row_count):
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, "", lot, sales_order, customer_po, expire
                                    ])
                            else:
                                # LPN不满足拆分，原样一行
                                core_data.append([
                                    invoice_no, delivery, doc_date, carrier, hawb,
                                    raw_qty_str, item_num, raw_lpn_str, "", lot, sales_order, customer_po, expire
                                ])
                        # 情况2：Serial有值，原有多字段匹配逻辑
                        else:
                            # 场景1：LPN数量=QTY，Serial数量也等于QTY，一一对应拆分
                            if lpn_count == qty_val and serial_count == qty_val and qty_val > 0:
                                print_log(f"匹配规则1：LPN/Serial数量={lpn_count}=QTY{qty_val}，逐行匹配拆分")
                                for idx in range(qty_val):
                                    single_lpn = lpn_list[idx]
                                    single_serial = serial_list[idx]
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, single_serial, lot, sales_order, customer_po, expire
                                    ])
                            # 场景2：仅1个LPN、仅1个Serial，按QTY循环复制N行
                            elif lpn_count == 1 and serial_count == 1 and qty_val > 0:
                                split_row_count = qty_val
                                single_lpn = lpn_list[0]
                                single_serial = serial_list[0]
                                print_log(f"匹配规则2：单LPN+单Serial，QTY={qty_val}，生成{split_row_count}条重复数据")
                                for _ in range(split_row_count):
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, single_serial, lot, sales_order, customer_po, expire
                                    ])
                            # 场景3：只有Serial多个、LPN单个，且serial_count == qty_val
                            elif lpn_count == 1 and serial_count == qty_val and qty_val > 0:
                                single_lpn = lpn_list[0]
                                print_log(f"匹配规则3：单LPN，Serial数量={serial_count}=QTY{qty_val}，拆分Serial多行")
                                for single_serial in serial_list:
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, single_serial, lot, sales_order, customer_po, expire
                                    ])
                            # 场景4：只有LPN多个、Serial单个，且lpn_count == qty_val
                            elif serial_count == 1 and lpn_count == qty_val and qty_val > 0:
                                single_serial = serial_list[0]
                                print_log(f"匹配规则4：单Serial，LPN数量={lpn_count}=QTY{qty_val}，拆分LPN多行")
                                for single_lpn in lpn_list:
                                    core_data.append([
                                        invoice_no, delivery, doc_date, carrier, hawb,
                                        "1", item_num, single_lpn, single_serial, lot, sales_order, customer_po, expire
                                    ])
                            # 其他所有不匹配场景，保留原始一行，LPN/Serial逗号拼接不拆分
                            else:
                                core_data.append([
                                    invoice_no, delivery, doc_date, carrier, hawb,
                                    raw_qty_str, item_num, raw_lpn_str, raw_serial_str, lot, sales_order, customer_po, expire
                                ])

                    print_log(f"{filename} 处理完成后总明细行数:{len(core_data)}")

                status = "成功"
                success_count += 1
                print_log(f"✅ [{filename}] 处理成功")

            else:
                raise Exception("OCR返回识别状态异常")

        except Exception as e:
            status = "失败"
            res_msg = str(e)
            fail_count += 1
            print_log(f"❌ [{filename}] 处理失败：{e}")

        full_log_data.append([filename, file_id, file_url, req_uuid, "", status, res_msg, ""])

    print_log(f"\n===== 批量处理结束 =====")
    print_log(f"总文件：{total_files} | 成功：{success_count} | 失败：{fail_count} | 有效明细行数：{len(core_data)}")

    # 写入飞书多维表格
    print_log("开始同步统计数据至飞书多维表格...")
    token = get_tenant_access_token()
    send_to_bitable(token, select_text, total_files)

    try:
        output_file = export_excel(select_text, core_data, full_log_data)
    except PermissionError:
        print_log("无法写入文件，请关闭 Excel 文件后重试。")
        ui_message_queue.put(("error", "无法写入文件，请关闭 Excel 文件后重试。"))
        return
    except Exception as e:
        print_log(f"Excel导出失败：{str(e)}")
        ui_message_queue.put(("error", f"Excel导出失败: {str(e)}"))
        return

    msg = (f"处理完成！\n成功文件数: {success_count}\n失败文件数: {fail_count}\n"
           f"拆分LPN/Serial后有效数据行数: {len(core_data)}\n"
           f"结果已导出至: {os.path.basename(output_file)}")
    ui_message_queue.put(("complete", msg))


def export_excel(select_text, core_data, full_log_data):
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(os.getcwd(), f"GE单据OCR识别结果_{timestamp}.xlsx")

    wb = Workbook()
    ws_core = wb.active
    ws_core.title = "识别结果列表"

    if select_text == "GE-ORACLE拣货单":
        # Order Number后面紧跟全部新增字段
        core_headers = [
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
        core_headers = [
            "服务申请号", "SR编号", "供应商", "SSO", "收货地址",
            "时效", "收货人电话", "申请说明",
            "收货人", "姓名", "客户设备id",
            "物料编号", "数量", "序列号", "跟踪号", "货位", "状态", "仓库"
        ]
    elif select_text == "GE-发票单":
        # DELIVERY后面紧跟 DATE、CARRIER、HAWB
        core_headers = [
            "INVOICE NO", "DELIVERY",
            "DATE", "CARRIER", "HAWB",
            "QTY", "ITEM NUMBER", "LPN Number", "Serial Number",
            "LOT Number", "SALES ORDER NO", "CUSTOMER PO", "Expiration Date"
        ]

    ws_core.append(core_headers)
    for row in core_data:
        ws_core.append(row)

    # 自适应列宽
    for col in ws_core.columns:
        max_len = max(len(str(cell.value)) for cell in col)
        ws_core.column_dimensions[col[0].column_letter].width = min(max_len + 2, 80)

    # 日志Sheet
    ws_log = wb.create_sheet(title="全量处理日志")
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


def poll_ui_queue():
    flush_log()
    while True:
        try:
            kind, payload = ui_message_queue.get_nowait()
        except queue.Empty:
            break
        if kind == "complete":
            btn.config(text="选择文件并开始处理", state=tk.NORMAL)
            messagebox.showinfo("完成", payload)
        elif kind == "error":
            btn.config(text="选择文件并开始处理", state=tk.NORMAL)
            messagebox.showerror("错误", payload)

    if worker_thread is not None and worker_thread.is_alive():
        win.after(100, poll_ui_queue)

# ========== 界面部分 ==========
win = tk.Tk()
win.title("GE单据批量OCR处理工具")
win.geometry("980x440")

tk.Label(win, text="选择模版规则：", font=("黑体", 11)).place(x=30, y=15)
combo_model = ttk.Combobox(win, width=28, font=("黑体", 11), state="readonly")
combo_model["values"] = list(MODEL_MAP.keys())
combo_model.set("")
combo_model.place(x=130, y=15)

tk.Label(win, text="支持格式: 图片 / PDF\n支持单据: GE‑ORACLE拣货单、GE‑OSCAR拣货单、GE‑发票单",
          font=("黑体", 12), justify=tk.LEFT).pack(pady=50)

btn = tk.Button(win, text="选择文件并开始处理", command=run_task, width=25, height=2,
                bg="#4CAF50", fg="white", font=("黑体", 12))
btn.pack(pady=10)

log_text = tk.Text(win, height=12, width=110)
log_text.pack(pady=10)

win.mainloop()
