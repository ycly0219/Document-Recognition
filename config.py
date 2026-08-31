"""集中管理 OCR、飞书接口与模板映射配置。"""

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
OCR_RETRY_INTERVAL = 5
OCR_MAX_POLL_SECONDS = 300

# 模板映射（GE拣货单、GE‑OSCAR拣货单、GE‑发票单）
MODEL_MAP = {
    "GE-ORACLE拣货单": "logistics_east_ge_picklist_99999_1503",
    "GE-OSCAR拣货单": "logistics_ge_oscarpicklist_1503",
    "GE-发票单": "logistics_east_ge_invoice_99999_1503"
}

# ---------------- Flux WMS putPurchaseOrder 对接配置 ----------------
WMS_PUT_PURCHASE_ORDER_URL = (
    "https://sinoewms-qas.i.sinotrans.com/datahubjson/FluxWmsJsonApi_WJC/"
    "?method=putPurchaseOrder&apptoken=0B61B741BB1970A66A63DD653A131D68"
    "&sign=123&format=json"
)
WMS_CUSTOMER_ID = "GEHC"
WMS_WAREHOUSE_ID = "WH004078"

# ---------------- 飞书多维表格配置 ----------------
FEISHU_APP_ID = "cli_aa978beae8f81cca"
FEISHU_APP_SECRET = "ywHBY0AmJc00TojMIghLzgRHpyngHXpR"
BITABLE_RECORDS_URL = "https://open.feishu.cn/open-apis/bitable/v1/apps/Jqfwbt2bWaLq7AswNz2c1iSUn4i/tables/tblQRssVXgCA7mEs/records"
# =============================================================
