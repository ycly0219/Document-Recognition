"""OCR 文件上传、任务提交与结果轮询。"""

import json
import os
import time

import requests

from config import (
    API_URL,
    GET_RESULT_API_URL,
    OCR_APP_ID,
    OCR_APP_KEY,
    OCR_APP_SECRET,
    OCR_CALLBACK_URL,
    OCR_DOC_TYPE,
    OCR_MAX_RETRY,
    OCR_ORG_ID,
    OCR_REGIONAL_CODE,
    OCR_RETRY_INTERVAL,
    OCR_SYS_CODE,
    OCR_SYS_CODE_OSCAR,
    ORG_ID,
    SOURCE_CODE,
    UPLOAD_API_URL,
)
from logging_utils import print_log


def get_file_type_and_content_type(file_path):
    """根据扩展名返回上传接口使用的文件类型与 Content-Type。"""
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
    """上传本地文件，返回 fileId 与文件路径。"""
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
    """提交异步 OCR 任务，返回 reqUuid。"""
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
    """按 reqUuid 轮询 OCR 识别结果，直到有有效 commitResult。"""
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
