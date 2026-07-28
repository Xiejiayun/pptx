# OOXML 演示格式识别设计

日期：2026-07-28  
状态：已批准实施

## 目标

为 `.pptx`、`.pptm`、`.ppsx`、`.ppsm`、`.potx`、`.potm` 建立稳定的公共类型模型，并在打开 package 时依据 presentation part 的 content type 识别格式。该能力是后续从零创建、显式格式转换、宏处理和六格式验证的共同基础。

本小项只负责格式建模和识别，不实现格式转换、宏删除、签名诊断或浏览器 I/O。

## 公共 API

在 model 层新增：

- `PresentationFormat`：六种受支持格式的字符串联合类型。
- `PresentationFormatProfile`：包含格式、扩展名、presentation content type、是否允许宏、是否为放映格式、是否为模板格式。
- `PRESENTATION_FORMAT_PROFILES`：只读格式表。
- `presentationFormatProfile(format)`：按格式取得 profile。
- `detectPresentationFormat(contentType)`：按 presentation content type 识别格式；无法识别时返回 `undefined`。
- `UnsupportedPresentationFormatError`：presentation part 存在但 content type 不受支持时抛出，携带实际 content type 和 part URI。

`PresentationModel` 新增只读 `format` 与 `formatProfile`。构造时只读取 root `officeDocument` relationship 指向的 presentation part，不根据文件扩展名猜测格式。

## 数据和兼容策略

格式映射使用 ECMA-376/Microsoft OOXML 的六个 presentation content type：普通、宏启用、放映、宏启用放映、模板、宏启用模板。打开后保持原格式；本小项不会修改主 part content type，也不会删除 VBA、签名或 opaque parts。

未知 presentation content type 不降级为 `.pptx`，避免调用方在错误假设下写回。旧二进制 `.ppt` 不是 ZIP/OPC package，由现有 package 打开流程拒绝。

## 错误处理

- 缺少 root presentation part：继续使用现有 `PackageError`。
- root part content type 不属于六种格式：抛出 `UnsupportedPresentationFormatError`。
- 文件名扩展名与 package content type 不一致：以 package content type 为准；文件名一致性诊断留给后续 validator 小项。

## 测试与验收

1. 六个 content type 均能打开并得到准确的 `format` 和 profile 标志。
2. 未知 presentation content type 明确失败，错误携带 part URI 与 content type。
3. `.pptx` 现有打开、编辑和无修改字节级 round-trip 测试继续通过。
4. 公共聚合包能导出新增类型、profile 表、检测函数和错误类型。

## 后续依赖

后续 `Presentation.create()` 直接以 profile 生成正确的 content type；格式转换 API 使用同一张表；宏/签名验证根据 profile 决定规则，不再散落字符串判断。
