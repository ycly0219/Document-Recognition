# Flux WMS `putOriginalSalesOrder` 字段映射

本文档约定 `GE-ORACLE拣货单` / `GE-OSCAR拣货单` 导出的销售订单 Excel 字段与 Flux WMS `putOriginalSalesOrder` 报文字段的映射关系。当前仅维护映射文档，不实现接口发送逻辑。

## 固定值

```text
FIXED_CUSTOMER_ID  = "GEHC"
FIXED_WAREHOUSE_ID = "WH004078"
CONSIGNEE_ID       = "CONSIGNEEID"
PACK_UOM           = "EA"
PRICE              = "0"
```

接口键名统一使用贴出的 JSON 样例键名；UE 侧示例中的 `loAtt` 按接口样例统一为 `lotAtt`。

## 头字段映射

| 接口字段 | ORACLE 来源 | OSCAR 来源 | 转换 / 固定值 |
| --- | --- | --- | --- |
| `warehouseId` | 固定 `WH004078`（A） | 固定 `WH004078`（A） | `FIXED_WAREHOUSE_ID` |
| `customerId` | 固定 `GEHC`（F） | 固定 `GEHC`（F） | `FIXED_CUSTOMER_ID` |
| `orderType` | 订单类型对应代码（C） | 订单类型对应代码（C） | `GNCK_*` / `GWCK_*` |
| `docNo` | `Order Number`（G） | `服务申请号`（G） | 原值 |
| `soReferenceA` | `Order Number`（G） | `服务申请号`（G） | 与 `docNo` 同源 |
| `soReferenceB` | 无来源省略 | `SR编号`（H） | 原值 |
| `soReferenceC` | `System Id`（I） | `客户设备id`（I） | 原值 |
| `soReferenceD` | 无来源省略 | 无来源省略 | 不发送 |
| `orderTime` | `Pick Slip Print Date`（D） | 导出当前时间（D） | ORACLE 转 `YYYY-MM-DD HH:MM:SS`；OSCAR 取导出时间 |
| `consigneeId` | 固定 `CONSIGNEEID`（V） | 固定 `CONSIGNEEID`（Z） | `CONSIGNEE_ID` |
| `consigneeAddress1` | `Ship To Address`（W） | `收货地址`（AA） | 原值 |
| `notes` | 当前无来源省略 | `申请说明`（AB） | 原值 |
| `hedi01` | `OrderType`（L） | `客户设备id`（I） | 与 `soReferenceC` 的 OSCAR 来源一致 |
| `hedi02` | `Ordered Date`（M） | 无来源省略 | 转 `YYYY-MM-DD` |
| `hedi03` | `Shipment Priority`（N） | 无来源省略 | 原值 |
| `hedi04` | `Ship Method`（O） | 无来源省略 | 原值 |
| `hedi05` | `Service LeveL`（P） | `时效`（P） | 原值 |
| `hedi06` | `FE SSO`（Q） | `SSO`（Q） | 原值 |
| `hedi07` | `FE Name`（R） | `姓名`（R） | 原值 |
| `hedi08` | `Shipping Instruction`（S） | 无来源省略 | 原值 |
| `hedi11` | `Special Instruction`（T） | 无来源省略 | 原值 |
| `hedi12` | `Pick From Subinv`（U） | 无来源省略 | 原值 |
| `hedi13` | `Pick From Subinv`（U） | `供应商`（V） | 与 `hedi12` 的 ORACLE 来源一致 |
| `hedi14` | 无来源省略 | `收货人`（W） | 原值 |
| `hedi15` | 无来源省略 | `收货人电话`（X） | 原值 |
| `userDefine1..5` | `udf01..udf05`（Y-AC） | `udf01..udf05`（AC-AG） | 列位置映射；ORACLE `udf01` 当前来自 `SHIP TO NO`，其余当前为空 |

## 明细字段映射

| 接口字段 | ORACLE 来源 | OSCAR 来源 | 转换 / 固定值 |
| --- | --- | --- | --- |
| `lineNo` | 发送行号 | 发送行号 | 从 `1` 开始递增 |
| `sku` | `Item Number`（AG） | `物料编号`（AK） | 原值 |
| `orderedQty` | `Qty`（AO） | `数量`（AS） | 原值 |
| `packUom` | 固定 `EA`（AP） | 固定 `EA`（AT） | `PACK_UOM` |
| `price` | 固定 `0`（AV） | 固定 `0`（AZ） | `PRICE` |
| `lotAtt04` | `Lot`（AI） | 无来源省略 | 原值 |
| `lotAtt05` | 固定 `ORACLE`（AJ） | 固定 `OSCAR`（AN） | 货物来源 |
| `lotAtt07` | `Org`（AK） | `仓库`（AO） | 原值 |
| `lotAtt08` | `Pick From Subinv`（AL） | 状态（AP） | ORACLE 按 `GD/BAD` 后缀转 `GOOD/BAD`；OSCAR `好件` 转 `GOOD`，其余留空 |
| `lotAtt09` | `Serial`（AM） | `序列号`（AQ） | 原值 |
| `lotAtt11` | `LPN`（AN） | 无来源省略 | 原值 |

## 不发送 / 未映射字段

- 接口无来源字段一律省略，不填空字符串或假值。
- 接口样例没有 `orderStatus`、`releaseStatus`、`lineStatus`，即使导出 Excel 有固定值，也不写入报文。
- 明细不发送 `customerId`，严格按确认后的接口 JSON 结构处理。
- `lotAtt02/15/16/17` 属于 `GE-发票单` 来源字段，Oracle/Oscar 销售订单导出 Excel 没有这些字段，不误搬进本映射。
- 头部无来源字段：`priority`、`expectedShipmentTime1`、`requiredDeliveryTime`、`consigneeName`、`consigneeContact`、`consigneeAddress2/3`、收货人国家/省/市/区/街道/邮箱/电话/邮编、承运商、开单方、结算方、`channel`、`hedi09/10`、`invoicePrintFlag`、`route`、`stop`、`createSource`、`transportation` 均不发送。
- 明细无来源字段：`referenceNo`、`lotAtt01/02/03/06/10/12-24`、`dedi04-20`、`userDefine1-6`、`notes` 均不发送。
- ORACLE `Task Id`（AW）、`Pick From Locator`（AY）在 Excel 模板对应 `dedi01/03`；OSCAR `跟踪号`（BB）、`货位`（BC）对应 `dedi02/03`。接口样例只给出 `dedi04-20`，序号和语义不匹配，不自动硬映射。

## 实现要点

- `订单类型` 中文标签应先转换为当前模板使用的 `GNCK_*` / `GWCK_*` 代码，再写入 `orderType`。
- ORACLE 日期转换可复用导出逻辑中的三字母/完整月份、2 位/4 位年份转换规则。
- `lotAtt08` 的 ORACLE 和 OSCAR 转换规则不同，不能使用同一个固定值。
- 报文仅包含有可靠来源或固定值的字段；未映射字段不拼接进 JSON。

