# Presentation RTL 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 presentation 根级从右到左模式增加原生创建、读取、编辑和清除能力，对等 PptxGenJS 4.0.1 的 `pptx.rtlMode` 有效输出，并覆盖本库支持的全部六种 presentation format。

本小项只处理 direct `p:presentation@rtl`。不修改 paragraph `a:pPr@rtl`、text body `a:bodyPr@rtlCol`、master/layout/default text style，不自动改变现有或后续文本的 alignment、run 顺序或 paragraph RTL。presentation RTL 与上一小项的 paragraph RTL 是两个独立层级。

## API

```ts
interface CreatePresentationOptions {
  readonly rtlMode?: boolean;
}

class PresentationModel {
  get rtlMode(): boolean | undefined;
  set rtlMode(value: boolean | undefined);
}

const document = PptxDocument.create({ rtlMode: true });
document.rtlMode = false;
document.rtlMode = undefined;
```

`CreatePresentationOptions.rtlMode` 负责 zero-input 创建；`PresentationModel.rtlMode` 同时服务 `PptxDocument` 与 model 层打开后的编辑。getter 是 direct-state snapshot，而不是 effective direction 推断。

true 写 `rtl="1"`，false 写 `rtl="0"`，undefined 清除 direct attribute。创建时省略字段不写属性；显式 false 写 `0`，保留调用者的直接 LTR 意图。PptxGenJS 4.0.1 的 false 会省略属性，但 presentation root 没有上层 direction inheritance，因此两者 effective behavior 相同；显式 `0` 是本库的 lossless editing 扩展。

## 方案选择

考虑过三种方案：

1. 创建选项与可编辑属性都使用 `rtlMode`；采用此方案。它保持 PptxGenJS 命名，覆盖从零创建与已有文件编辑，并提供 direct 三态。
2. 只提供 `document.rtlMode` setter。表面更小，但从零创建需要先建包再变更，无法在 create option 中严格验证，也不符合现有 slide-size/format 创建配置模式。
3. 使用 `direction: 'ltr' | 'rtl'`。语义直观，但会偏离 PptxGenJS 公共名称，并丢失 direct attribute 缺失状态。

不复用 paragraph `RichTextParagraph.rtl` 作为全局默认，也不在 `AddTextOptions` 中读取 presentation state。调用者要让具体 paragraph 生效，仍需显式使用 paragraph `rtlMode` / `rtl`。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `pptx.rtlMode = true` 在 `p:presentation` 根写 `rtl="1"`。
- false 或省略不写 root `rtl`。
- truthy 非 boolean 值同样写 `rtl="1"`；runtime 不校验类型。
- 该属性不改变 `defaultTextStyle` 中各级 `a:lvlNpPr@rtl="0"`，也不替每个 paragraph 写 direction。

本库匹配 boolean true 的有效输出，严格拒绝非 boolean，并读取/编辑既有 direct root state。adapter 继续只消费标准 OOXML，不访问 PptxGenJS 私有字段。

## OOXML 映射与读取

```xml
<p:presentation
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  rtl="1">
  ...
</p:presentation>
```

getter 只读取根 `p:presentation` 的 direct `rtl`：`1`、`true`、`on` 返回 true，`0`、`false`、`off` 返回 false；缺失、空值或未知 token 返回 undefined。读取不产生 package mutation。

setter 通过 lossless source-span patch 仅替换、插入或删除 root attribute。替换保留 attribute 名称、位置和所有其他 root attributes；插入放在 start tag 结束前；删除连同紧邻的水平空白移除。未知 token 在无关 mutation 中原样保留，在显式 setter 时按 true/false/undefined 规则替换或删除。

## 创建、编辑与事务边界

创建入口在创建任何 package part 前验证 `rtlMode`。true/false 直接渲染到 canonical presentation XML；undefined 保持当前 root 输出。format、slide size、master/layout/theme、notes、relationships 与 default text style 不受影响。

编辑 setter 在 `OpcPackage.transaction()` 内运行。值只能是 boolean 或 undefined；null、number、string、object、array、symbol 都在 mutation 前抛出 TypeError。失败必须保持 presentation bytes、mutation journal、slides/models identity 与其他 package parts 不变。

设置同一值仍允许 canonicalize recognized/unknown lexical form，但只产生 presentation part mutation。transaction rollback 恢复 exact bytes 与 journal。write/reopen 保留 true、false 和 cleared 状态。

## 文档与兼容边界

兼容表新增 presentation `pptx.rtlMode` 已支持行，并明确与 paragraph RTL 分离。API 与 package README 展示 create option 和 editable property。changelog 记录 strict direct root support。

本小项不宣称 PowerPoint/LibreOffice 会自动重排所有已存在的 shape；它只保证标准 OOXML root flag 的正确序列化、读取和保持。视觉验证使用同一内容的 PptxGenJS true 基准，检查文件无修复打开、无裁切，并以 XML/CLI validation 为主要 conformance 证据。

## 测试与验收

1. model strict getter 覆盖六种 boolean lexical form、缺失、空和未知 token，且只读不产生 mutation。
2. setter 覆盖 true、false、clear、unknown-token replacement、unrelated root attributes/children 保留、rollback 与 stable identity。
3. native create 覆盖 omitted/true/false、全部六种格式、invalid input 在 package 返回前拒绝、write/reopen 与 validator。
4. PptxGenJS conformance 覆盖 true、false/omitted controls、truthy invalid baseline，并确认 paragraph/defaultTextStyle 不被误读为 root state。
5. packed Node/browser/declaration smoke 覆盖 create、getter、setter 与 clear。
6. 全仓 typecheck/test、独立 performance、actual tarball smoke、CLI PowerPoint 2010 profile validation 与 LibreOffice export 全部通过。
