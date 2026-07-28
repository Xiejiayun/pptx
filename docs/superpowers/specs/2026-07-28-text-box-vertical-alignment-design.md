# 文本框垂直对齐设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为文本 shape 增加 top、middle、bottom 三种文本框垂直对齐的原生创建、读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.valign` 的有效行为。该能力属于 text body 布局，映射到 direct `a:bodyPr@anchor`；它不属于 paragraph alignment，也不改变现有 margin、bullet、spacing、run style 或 shape transform。

本小项只包含普通文本 shape 的 direct vertical anchor。不包含 table-cell valign、placeholder/master inheritance 编辑、OOXML `just`/`dist` 创建、`anchorCtr`、autofit、wrap、text direction、RTL columns、column count/spacing 或 paragraph vertical metrics。这些能力继续拆成独立小项；读取和非 valign mutation 不得破坏其原始 XML。

## API

```ts
export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';

export interface AddTextOptions {
  readonly valign?: TextBoxVerticalAlignment;
}

export class ShapeModel {
  get verticalAlignment(): TextBoxVerticalAlignment | undefined;
  set verticalAlignment(value: TextBoxVerticalAlignment | undefined);
}
```

创建时沿用 PptxGenJS 的 `valign` 名称，便于迁移：

```ts
const shape = slide.addText('Bottom aligned', {
  x: inches(1),
  y: inches(1),
  width: inches(4),
  height: inches(1),
  valign: 'bottom',
});

shape.verticalAlignment = 'top';
shape.verticalAlignment = undefined; // remove the direct anchor
```

`addText()` 和 `addRichText()` 共用相同 text-body 创建路径。`valign` 省略时继续显式写 middle，保持本库既有 `anchor="ctr"` 默认并匹配 PptxGenJS 普通文本框 runtime。

`shape.verticalAlignment` getter 只读取当前 shape 的 direct `a:bodyPr@anchor`：`t`、`ctr`、`b` 分别返回 top、middle、bottom；缺少属性或 token 不在这三个值时返回 `undefined`。setter 写入所选 direct token；赋值 `undefined` 移除 direct `anchor`，不伪造默认值或继承值。

## 方案选择

考虑过三种 API/实现方案：

1. 创建和编辑都使用 `valign`。它与 PptxGenJS 最接近，但 live shape 上缩写不如 `verticalAlignment` 清楚，也容易与 table-cell 属性混淆。
2. 创建和编辑都使用 `verticalAlignment`。命名统一，但会增加现有 PptxGenJS 调用迁移成本，并与已采用的 `align`/`valign` creation convention 不一致。
3. `AddTextOptions.valign` 对齐 PptxGenJS，`ShapeModel.verticalAlignment` 提供清楚的 live property；采用此方案。

实现边界考虑过：

1. 在 vertical-alignment codec 复制 direct `txBody/bodyPr` 定位和 source-span attribute patch。代码最局部，但会重复 margin 已验证的错误处理与 self-closing 逻辑。
2. 把垂直对齐塞进 `text-box-margins.internal.ts`。改动少，但文件职责和名称都会失真。
3. 抽取窄 `text-body-properties.internal.ts`，只提供 direct `bodyPr` 定位与单 attribute lossless update；margin 和 vertical alignment 保持各自的 normalize/read/render/replace codec。它消除当前真实重复，同时不抽象 fit child、inheritance 或所有 bodyPr schema，因此采用。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- 普通文本框省略 `valign` 时写 `anchor="ctr"`。
- `valign: 'top'` 写 `anchor="t"`。
- `valign: 'middle'` 写 `anchor="ctr"`。
- `valign: 'bottom'` 写 `anchor="b"`。
- `TextBaseProps` 注释称默认 top，`TextPropsOptions` 注释称默认 middle；4.0.1 普通 `addText()` runtime 实际为 middle。
- 非法字符串不会报错，而是保留初始化的 `anchor="ctr"`。
- `TextProps[]` 单个 run 的 `options.valign` 不影响 text body；只有文本框 options 生效。

本库以普通 `addText()` 的有效 runtime 为 conformance 基线：省略时 middle，三种合法值映射一致。不复制非法值静默回退缺陷，也不暴露无效的 run-level vertical alignment。

## OOXML 映射与严格读取

```xml
<a:bodyPr wrap="square" rtlCol="0" anchor="b"/>
```

| 公共值 | `a:bodyPr@anchor` |
| --- | --- |
| `top` | `t` |
| `middle` | `ctr` |
| `bottom` | `b` |

OOXML 还允许 `just` 与 `dist`。它们不在 PptxGenJS `VAlign` 公共范围内，本小项 getter 不把它们错误归类为 top/middle/bottom；只读和无关 mutation 必须原样保留。空值、大小写变体、未知 token 和带空白 token 同样返回 `undefined`。

getter 不根据缺失 attribute 合成 schema default，也不读取 layout/master placeholder inheritance。这样 `undefined` 稳定表示“当前 shape 没有受支持的 direct anchor”。

## 编辑与 Lossless 边界

新建 text shape 继续保留 `wrap="square"`、margin attributes、`rtlCol="0"` 和其他既有 body properties，只把 `anchor` 设为目标 token。plain 与 rich 创建结果一致。

setter 定位 direct `txBody/bodyPr`，只增加、替换或移除 `anchor` attribute。既有 attribute 的 quote style/order、`anchorCtr`、margin、wrap、vert、numCol、spcCol、autofit child、extension、namespace 和未知内容保持原字节与相对顺序。self-closing `bodyPr` 增加 anchor 时仍保持合法 self-closing XML；缺少 direct `txBody` 或 `bodyPr` 时抛出 `ModelParseError`，不猜测或创建结构。

`.text`、`.richText`、`textMargins`、transform 和其他非 vertical-alignment mutation 保留原 anchor XML。设置 vertical alignment 不重建 paragraphs/runs，不改变 shape identity、relationships 或其他 package state。

## 验证与错误处理

输入必须精确为 `'top'`、`'middle'`、`'bottom'` 或 setter 的 `undefined`。其他 string、boolean、number、null、object、array 和 symbol 明确失败。创建 options 的 `undefined` 表示采用既有 middle default；live setter 的 `undefined` 表示移除 direct anchor。

所有输入在 package mutation 前完整验证。失败不得改变 part bytes、mutation journal、live `ShapeModel` identity、text、margin 或 vertical-alignment snapshot；外层 transaction rollback 同样完整恢复。

## 测试与验收

1. `addText()` 和 `addRichText()` 覆盖 omitted/middle/top/bottom，并验证与 margin、paragraph alignment、bullet、rich runs 的组合。
2. getter 严格读取 `t/ctr/b`；对 absent、`just`、`dist`、空值、大小写、前后空白与未知 token 返回 `undefined`，只读不产生 mutation。
3. `shape.verticalAlignment` 可在 self-closing/expanded `bodyPr` 上增加、更换和清除 anchor，保留 margins、`anchorCtr`、autofit/unknown children、paragraphs/runs 与 stable identity；write/reopen 和 duplicate 一致。
4. plain/rich text、margin 与 transform 编辑保留 anchor；缺失 `txBody/bodyPr`、非法输入及内外层 rollback 不改变 bytes、journal 或 live snapshots。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted/top/middle/bottom 和 ignored run-level `valign` 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和仓库 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 top/middle/bottom 对照页；用同版本 PptxGenJS 文件核对三种垂直位置和默认 middle。
