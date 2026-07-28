# Text Box Text Direction Design

## 目标与范围

为普通文本 shape 增加文本方向的原生创建、读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.vert` 的七个公开有效值。该能力映射到 direct `a:bodyPr@vert`，控制文本框内文字的横排、旋转竖排、东亚/蒙古文竖排与 WordArt 竖排方向；它不旋转 shape transform，也不改写字符、paragraph 或 run。

本小项只包含普通文本 shape 的 direct text-direction override。不包含 table-cell `textDirection`、placeholder/master inheritance 编辑、paragraph RTL、`rtlCol`、BiDi run properties、文字语言、autofit、overflow、columns 或 arbitrary shape rotation。这些能力继续拆成独立小项；读取和非 direction mutation 不得破坏其原始 OOXML。

## 公共 API

公共值与 PptxGenJS 4.0.1 的有效 `vert` union 一致：

```ts
export type TextBoxTextDirection =
  | 'eaVert'
  | 'horz'
  | 'mongolianVert'
  | 'vert'
  | 'vert270'
  | 'wordArtVert'
  | 'wordArtVertRtl';

interface AddTextOptions {
  readonly vert?: TextBoxTextDirection;
}
```

live shape 使用完整语义名：

```ts
class ShapeModel {
  get textDirection(): TextBoxTextDirection | undefined;
  set textDirection(value: TextBoxTextDirection | undefined);
}
```

用法：

```ts
const shape = slide.addText('Vertical label', {
  vert: 'vert270',
});

shape.textDirection = 'wordArtVert';
shape.textDirection = undefined; // remove the direct vert override
```

`addText()` 和 `addRichText()` 共用相同 text-body 创建路径。`vert` 省略时不写 direct attribute，保持本库既有输出并匹配 PptxGenJS 普通文本框 runtime。`shape.textDirection` getter 只读取 direct `bodyPr@vert`，setter 只更新或移除该 attribute；它不计算继承或把 absent 伪装为 `horz`。

## 方案选择

考虑过三个 API 方案：

1. 创建与 live 编辑都使用 `vert`。PptxGenJS 迁移最短，但 `shape.vert` 对非 OOXML 使用者不清楚，也容易被误解为 vertical alignment。
2. 创建与 live 编辑都使用 `textDirection`。语义统一，但偏离 PptxGenJS 的实际有效入口；4.0.1 同名 `textDirection` 在普通 slide text 上反而被忽略。
3. `AddTextOptions.vert` 对齐 PptxGenJS 有效入口，`ShapeModel.textDirection` 清楚表达 live property；采用此方案。它延续 `valign` / `verticalAlignment` 和 `wrap` / `textWrap` 的 creation/live 命名模式。

实现层采用新的窄 `text-box-text-direction.internal.ts` codec，复用 `requireTextBodyProperties()` 与 `updateTextBodyAttribute()`。不建立通用 bodyPr schema abstraction，因为 fit child、RTL、columns 与 overflow 各自有不同存储和清除语义，提前合并会模糊 ownership。

用户已授权实现方持续选择最佳方案并直接推进，因此本设计按迁移兼容、语义清晰和最小职责边界定稿。

## PptxGenJS 4.0.1 对照

使用公开 `addText()` / `write()` 实测：

- 省略 `vert` 不写 `bodyPr@vert`。
- `eaVert`、`horz`、`mongolianVert`、`vert`、`vert270`、`wordArtVert`、`wordArtVertRtl` 都原样写入 direct `vert`。
- 4.0.1 对任意 truthy 非法字符串也原样写出；本库不复制该无效 OOXML 缺陷，只接受七个公开 token。
- 普通 text box options 上的 `textDirection` 被 4.0.1 runtime 忽略。
- `TextProps[]` 单个 run 的 `vert` / `textDirection` 都不影响 text body。
- table cell 的 `textDirection` 是另一个有效 API，序列化到 cell `bodyPr@vert`，留作 table-cell 小项。

原生 API 覆盖七个有效 `vert` 行为，不暴露普通文本框上无效的 alias，也不把 run-level direction 塞入 `RichTextRunStyle`。adapter 继续只导入 PptxGenJS 的公开 OOXML 输出。

## OOXML 映射

创建旋转竖排文本的目标结构示例：

```xml
<a:bodyPr wrap="square" rtlCol="0" anchor="ctr" vert="vert270"/>
```

七个公共值与 `a:bodyPr@vert` 一一同名映射，不做大小写转换或别名归一。getter 只接受完全一致、无前后空白的七个 token。absent、空值、case variant、带空白值与未知 token 返回 `undefined`；只读访问不增加 mutation。

`horz` 是显式 direct override，与 absent 不等价：前者 getter 返回 `horz`，后者返回 `undefined`。live setter 赋值 `undefined` 移除 direct override，不写默认值。

## 创建与无损编辑

`validateAddTextOptions()` 在任何 package mutation 前验证 `vert`。normalized creation options 持有可选 `textDirection`; 省略时保持 `undefined`。plain/rich 两条路径把它传到 `textShapeXml()`；只有存在值时才由 codec 生成 ` vert="..."`。

live setter 定位 direct `txBody/bodyPr`，只增加、替换或移除 `vert` attribute。既有 quote style/order、wrap、margins、`anchor`、`anchorCtr`、`rtlCol`、`numCol`、`spcCol`、autofit child、extension、namespace 与未知内容保持原字节和相对顺序。self-closing/expanded `bodyPr` 都保持合法；缺少 direct `txBody` 或 `bodyPr` 时抛出 `ModelParseError`，不创建推测结构。

`.text`、`.richText`、`textMargins`、`verticalAlignment`、`textWrap`、transform 和其他非 direction mutation 保留原 vert XML。设置 direction 不重建 paragraphs/runs，不改变 shape identity、relationships 或其他 package state。

## 输入失败与事务

创建和 live setter 对 null、boolean、number、空字符串、case/whitespace variant、OOXML 之外字符串、array、object 和 symbol 抛出 `TypeError`。所有 validation 在 package mutation 前完成。

失败不得改变 part bytes、mutation journal、live `ShapeModel` identity、text、margin、vertical-alignment、wrapping 或 direction snapshot；外层 transaction rollback 同样完整恢复。

## 测试与发布门禁

验收覆盖：

1. `addText()` / `addRichText()` 对 omitted 和七个有效 token 生成正确 direct attribute，并与 margin、vertical alignment、wrapping、paragraph/run style 组合。
2. getter 严格读取七个 direct token；absent、空值、case variant、空白和未知 token 返回 `undefined`，只读不产生 mutation。
3. `shape.textDirection` 可在 self-closing/expanded `bodyPr` 上增加、更换和清除 vert，保留 fit/RTL/unknown metadata、paragraphs/runs 与 stable identity；write/reopen 和 duplicate 一致。
4. plain/rich text、margin、vertical alignment、wrapping 与 transform 编辑保留 vert；缺失 `txBody/bodyPr`、非法输入及 rollback 不改变 bytes、journal 或 live snapshots。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted、七个有效 token、invalid passthrough、ignored outer `textDirection` 与 ignored run-level direction 的真实输出导入。
6. Node/browser/declaration tarball smoke 覆盖创建、读取、编辑和 clear assignment。
7. 原生与 PptxGenJS 方向对照文件均通过 `powerpoint-2010` CLI profile 的 0 error / 0 warning 验证；LibreOffice 无修复导出，视觉核对至少覆盖 `horz`、`vert`、`vert270` 与 `wordArtVert` 的方向一致性，并记录客户端对其余 token 的降级情况。

完成后更新 changelog、公共 API、npm README 与兼容矩阵，把普通文本框 `vert` 从剩余 partial row 中移除，并明确 table-cell `textDirection` 仍未覆盖。全仓测试、独立性能、真实 tarball、CLI 和 LibreOffice 对照全部通过后，才允许 implementation commit 与 push。
