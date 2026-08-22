"""飞书 Token 获取与多维表格统计写入。"""

import requests

from config import BITABLE_RECORDS_URL, FEISHU_APP_ID, FEISHU_APP_SECRET
from logging_utils import print_log


def get_tenant_access_token():
    """获取飞书 tenant_access_token，失败时返回 None。"""
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
    """把本批统计写入飞书多维表格。"""
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
