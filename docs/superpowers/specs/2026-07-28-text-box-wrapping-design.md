# Text Box Wrapping Design

## 目标与范围

为普通文本 shape 增加文本框自动换行的原生创建、读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.wrap` 的有效行为。该能力属于 text body 布局，映射到 direct `a:bodyPr@wrap`；它不改变 paragraph、run、shape geometry 或文本内容。

本小项只包含普通文本 shape 的 direct wrapping override。不包含 table-cell wrapping、placeholder/master inheritance 编辑、overflow、autofit、vertical text、RTL columns、column count/spacing、word breaking 或手工插入换行。这些能力继续拆成独立小项；读取和非 wrapping mutation 不得破坏其原始 OOXML。

## 公共 API

创建 API 沿用 PptxGenJS 的选项名：

```ts
interface AddTextOptions {
  readonly wrap?: boolean;
}
```

live shape 使用语义更明确的属性名：

```ts
class ShapeModel {
  get textWrap(): boolean | undefined;
  set textWrap(value: boolean | undefined);
}
```

用法：

```ts
const shape = slide.addText('A long line of text', {
  width: inches(2),
  wrap: false,
});

shape.textWrap = true;
shape.textWrap = undefined; // remove the direct wrap override
```

`addText()` 和 `addRichText()` 共用相同 text-body 创建路径。`wrap` 省略时继续显式写 true，保持本库既有 `wrap="square"` 默认并匹配 PptxGenJS 普通文本框 runtime。

`shape.textWrap` getter 只读取当前 shape 的 direct `a:bodyPr@wrap`：`square` 返回 true，`none` 返回 false；缺少属性或 token 不受支持时返回 `undefined`。setter 写入所选 direct token；赋值 `undefined` 移除 direct `wrap`，不伪造 OOXML 默认值或继承值。

## 方案选择

考虑过三个公开表面：

1. 创建与 live 编辑都叫 `wrap`。迁移最短，但 `shape.wrap` 容易被理解为 shape/group 包装，而不是 text body 布局。
2. 创建与 live 编辑都叫 `textWrap`。命名统一，但会增加 PptxGenJS 调用迁移成本。
3. `AddTextOptions.wrap` 对齐 PptxGenJS，`ShapeModel.textWrap` 明确表达 live text-body property；采用此方案。它与已经采用的 `valign` / `verticalAlignment` 非对称模式一致。

实现层考虑过：

1. 在 `slide.ts` 直接拼接和读取 token。代码少，但会把 OOXML codec 细节继续堆入创建模型。
2. 建立通用 bodyPr schema abstraction。它能覆盖未来 fit、vert、RTL 和 columns，但会在本小项中过度建模尚未确认的语义。
3. 新建窄 `text-box-wrapping.internal.ts`，复用已经完成的 `requireTextBodyProperties()` 与 `updateTextBodyAttribute()`，只负责 boolean/token normalize、render、read 和 replace；采用此方案。

用户已明确授权由实现方全权决定最佳方案并持续推进，因此本设计按上述可迁移性、语义清晰度与最小职责边界直接定稿，不再设置逐项询问门槛。

## PptxGenJS 4.0.1 对照

使用公开 `addText()` / `write()` 实测：

- 省略 `wrap` 写 `wrap="square"`。
- `wrap: true` 写 `wrap="square"`。
- `wrap: false` 写 `wrap="none"`。
- 非 boolean 值被 runtime 静默回退到 `square`；本库不复制该宽松缺陷，非法输入明确失败。
- `TextProps[]` 单个 run 的 `options.wrap` 不影响 text body；本库不在 run style 上暴露 wrap。

原生 API 对前三项有效行为保持一致。adapter 继续只导入 PptxGenJS 的公开 OOXML 输出，不读取私有字段。

## OOXML 映射

创建关闭自动换行的目标结构：

```xml
<a:bodyPr wrap="none" rtlCol="0" anchor="ctr"/>
```

映射固定为：

| 公共值 | `a:bodyPr@wrap` |
| --- | --- |
| `true` | `square` |
| `false` | `none` |

`square` 是 DrawingML 文本框启用自动换行时的标准 token；这里的公共 boolean 不暴露 OOXML 命名细节。

getter 只接受大小写完全一致、无前后空白的 direct `square` / `none`。缺少属性、空值、case variant、带空白值和未知 token 都返回 `undefined`。只读访问不得增加 mutation。

## 创建与无损编辑

`validateAddTextOptions()` 在任何 package mutation 前验证 `wrap`：只有 boolean 或省略合法。normalized creation options 持有必需的 `textWrap: boolean`；省略值归一为 true。plain/rich 两条路径把归一值传到 `textShapeXml()`，由 wrapping codec 生成单个 attribute。

live setter 定位 direct `txBody/bodyPr`，只增加、替换或移除 `wrap` attribute。既有 attribute 的 quote style/order、margin、`anchor`、`anchorCtr`、`vert`、`rtlCol`、`numCol`、`spcCol`、autofit child、extension、namespace 和未知内容保持原字节与相对顺序。self-closing 与 expanded `bodyPr` 都必须保持合法；缺少 direct `txBody` 或 `bodyPr` 时抛出 `ModelParseError`，不猜测或创建结构。

`.text`、`.richText`、`textMargins`、`verticalAlignment`、transform 和其他非 wrapping mutation 保留原 wrap XML。设置 wrapping 不重建 paragraphs/runs，不改变 shape identity、relationships 或其他 package state。

## 输入失败与事务

创建和 live setter 对 null、number、string、array、object 与 symbol 等非 boolean 输入抛出 `TypeError`。OOXML token `square` / `none` 也不是公共输入。

所有 validation 在 package mutation 前完成。失败不得改变 part bytes、mutation journal、live `ShapeModel` identity、text、margin、vertical-alignment 或 wrapping snapshot；外层 transaction rollback 同样完整恢复。

## 测试与发布门禁

验收覆盖：

1. `addText()` / `addRichText()` 的 omitted、true、false 分别写 `square`、`square`、`none`，并与 margin、vertical alignment、paragraph/run style 组合。
2. getter 严格读取 direct `square` / `none`；absent、空值、case variant、空白和未知 token 返回 `undefined`，只读不产生 mutation。
3. `shape.textWrap` 可在 self-closing/expanded `bodyPr` 上增加、更换和清除 wrap，保留 fit/vert/RTL/unknown metadata、paragraphs/runs 与 stable identity；write/reopen 和 duplicate 一致。
4. plain/rich text、margin、vertical alignment 与 transform 编辑保留 wrap；缺失 `txBody/bodyPr`、非法输入及内外层 rollback 不改变 bytes、journal 或 live snapshots。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted/true/false、invalid fallback 和 ignored run-level wrap 的真实输出导入。
6. Node/browser/declaration tarball smoke 覆盖创建、读取、编辑和 clear assignment。
7. 原生与 PptxGenJS 长文本对照文件均通过 `powerpoint-2010` CLI profile 的 0 error / 0 warning 验证；LibreOffice 无修复导出，并在窄文本框中显示 wrap true 与 false 的明确差异。

完成后更新 changelog、公共 API、npm README 与兼容矩阵，把 `wrap` 从剩余 text-body partial row 中移除。全仓测试、独立性能门禁、真实 tarball、CLI 和 LibreOffice 对照全部通过后，才允许 implementation commit 与 push。
