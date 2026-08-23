"""演示用固定样例数据，结构与真实解析结果一致，不调用接口。"""

from parsers import get_core_headers


def generate_mock_data(select_text):
    """按模板生成两个演示文件页签的模拟明细数据。"""
    headers = get_core_headers(select_text)
    if select_text == "GE-ORACLE拣货单":
        rows = [
            ["PO240821-001", "T1001", "ITEM-A001", "10", "LPN-A18-480056",
             "SN-12444", "", "CN", "A-01-01",
             "ORACLE", "19-MAY-26", "P1", "AIR", "STANDARD",
             "SSO001", "ZHANGSAN", "GE HEALTHCARE", "CUS-1001",
             "NO.100 ZHANGJIANG ROAD SHANGHAI", "FE@GE.COM",
             "SHIP ASAP", "", "99999", "19-MAY-26 16:44:22", "SYS-01",
             "SUBINV-A", "PO-2026-001", "DEL-20260821-01"],
            ["PO240821-001", "T1002", "ITEM-B001", "5", "LPN-B02",
             "", "LOT-B02", "US", "B-02-03",
             "ORACLE", "19-MAY-26", "P1", "AIR", "STANDARD",
             "SSO001", "ZHANGSAN", "GE HEALTHCARE", "CUS-1001",
             "NO.100 ZHANGJIANG ROAD SHANGHAI", "FE@GE.COM",
             "SHIP ASAP", "", "99999", "19-MAY-26 16:44:22", "SYS-01",
             "SUBINV-A", "PO-2026-001", "DEL-20260821-01"],
            ["PO240821-002", "T2001", "ITEM-C001", "8", "LPN-C03",
             "SN-C03", "LOT-C03", "DE", "C-03-05",
             "ORACLE", "19-MAY-26", "P2", "SEA", "STANDARD",
             "SSO002", "LISI", "GE HEALTHCARE", "CUS-1002",
             "NO.200 PUJIAN ROAD SHANGHAI", "FE2@GE.COM",
             "", "", "99999", "19-MAY-26 16:44:22", "SYS-02",
             "SUBINV-B", "PO-2026-002", "DEL-20260821-02"]
        ]
    elif select_text == "GE-OSCAR拣货单":
        rows = [
            ["SA20260821-01", "SR-001", "GE SUPPLIER", "SSO-OSC-01",
             "NO.1 SUPPLY ROAD SHANGHAI", "24H", "13800000001", "APPLICATION NOTE",
             "RECEIVER-01", "LI SI", "DEVICE-001",
             "MT-OSC-001", "12", "SN-0001", "TR-0001", "LOC-01", "TO RECEIVE", "WH-A"],
            ["SA20260821-01", "SR-001", "GE SUPPLIER", "SSO-OSC-01",
             "NO.1 SUPPLY ROAD SHANGHAI", "24H", "13800000001", "APPLICATION NOTE",
             "RECEIVER-01", "LI SI", "DEVICE-001",
             "MT-OSC-002", "6", "SN-0002", "TR-0002", "LOC-02", "TO RECEIVE", "WH-A"],
            ["SA20260821-02", "SR-002", "GE SUPPLIER", "SSO-OSC-02",
             "NO.2 SUPPLY ROAD SHANGHAI", "48H", "13900000002", "APPLICATION NOTE 2",
             "RECEIVER-02", "WANG WU", "DEVICE-002",
             "MT-OSC-003", "3", "SN-0003", "TR-0003", "LOC-03", "PICKED", "WH-B"]
        ]
    elif select_text == "GE-发票单":
        # 前两行演示 LPN/Serial 按数量拆分后的效果
        rows = [
            ["INV20260821-01", "ITEM-INV-001", "1", "LPN-1", "", "LOT-01",
             "2026/09/30", "CN", "SO-001", "PO-001",
             "2026/08/21", "DEL-20260821", "FEDEX", "HAWB-001"],
            ["INV20260821-01", "ITEM-INV-001", "1", "LPN-2", "", "LOT-01",
             "2026/09/30", "CN", "SO-001", "PO-001",
             "2026/08/21", "DEL-20260821", "FEDEX", "HAWB-001"],
            ["INV20260821-01", "ITEM-INV-001", "1", "LPN-3", "", "LOT-01",
             "2026/09/30", "CN", "SO-001", "PO-001",
             "2026/08/21", "DEL-20260821", "FEDEX", "HAWB-001"],
            ["INV20260821-01", "ITEM-INV-002", "1", "LPN-4", "SN-A", "LOT-02",
             "2026/10/31", "US", "SO-002", "PO-002",
             "2026/08/21", "DEL-20260821", "FEDEX", "HAWB-001"],
            ["INV20260821-01", "ITEM-INV-002", "1", "LPN-4", "SN-B", "LOT-02",
             "2026/10/31", "US", "SO-002", "PO-002",
             "2026/08/21", "DEL-20260821", "FEDEX", "HAWB-001"],
            ["INV20260821-02", "ITEM-INV-003", "1", "LPN-9", "SN-C", "LOT-03",
             "2026/11/30", "DE", "SO-003", "PO-003",
             "2026/08/22", "DEL-20260822", "DHL", "HAWB-002"]
        ]
    else:
        raise ValueError(f"未知模板: {select_text}")

    split_index = max(1, len(rows) // 2)
    file_results = [
        {
            "filename": "模拟文件-01.png",
            "status": "成功",
            "message": "模拟数据，未调用 OCR",
            "rows": rows[:split_index],
        },
        {
            "filename": "模拟文件-02.pdf",
            "status": "成功",
            "message": "模拟数据，未调用 OCR",
            "rows": rows[split_index:],
        },
    ]
    return headers, file_results
