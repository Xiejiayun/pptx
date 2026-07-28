# 段落间距与行距设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为普通文本和结构化 rich text 增加段前、段后与行距，使原生创建和已有 PPT 编辑覆盖 PptxGenJS 4.0.1 的 `paraSpaceBefore`、`paraSpaceAfter`、`lineSpacing` 和 `lineSpacingMultiple`。能力按 paragraph 表达，并与已经支持的 alignment、bullet/numbering、rich runs、空 paragraph 和 soft break 正确组合。

本小项不包括普通段落左/右 margin、首行 indent、`indentLevel`、tab stop、文本框 margin/fit/wrap/vertical alignment，或未被 PptxGenJS 公开的段前/段后百分比值。这些继续作为独立小项；读取和非 rich-text mutation 不得破坏其原始 XML。

## API

```ts
type ParagraphLineSpacing =
  | {
      readonly kind: 'exact';
      readonly points: number;
    }
  | {
      readonly kind: 'multiple';
      readonly factor: number;
    };

interface ParagraphSpacing {
  readonly before?: number; // points
  readonly after?: number; // points
  readonly line?: ParagraphLineSpacing | false;
}

interface AddTextOptions extends Partial<Transform> {
  readonly spacing?: ParagraphSpacing;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly spacing?: ParagraphSpacing | false;
}

slide.addText('First\nSecond', {
  spacing: {
    before: 6,
    after: 8,
    line: { kind: 'multiple', factor: 1.5 },
  },
});

slide.addRichText(
  [
    { runs: [{ text: 'Uses defaults' }] },
    { runs: [{ text: 'Exact 24pt' }], spacing: { line: { kind: 'exact', points: 24 } } },
    { runs: [{ text: 'No spacing' }], spacing: false },
  ],
  { spacing: { before: 4, after: 6, line: { kind: 'multiple', factor: 1.2 } } },
);
```

`AddTextOptions.spacing` 是创建默认值。paragraph 未提供 `spacing` 时继承默认值；paragraph 对象按 `before`、`after`、`line` 三个字段逐项覆盖默认值，匹配 PptxGenJS 的预期继承语义。`spacing: false` 清除整组默认值，`before: 0`、`after: 0` 或 `line: false` 分别清除一个继承字段。

对已有 shape 设置 `richText` 时没有文本框默认值：未提供 `spacing` 表示目标 paragraph 不包含本地 spacing；对象中的未提供字段同样不写入，`spacing: false` 明确清除整组。getter 返回规范化对象，不返回 `false`，零值 spacing 规范化为字段缺失。

段前、段后和 exact line spacing 使用 point。它们的合法 OOXML 范围是 0–1584pt；exact line spacing 必须大于 0。multiple factor 使用 `1.5` 表示 1.5 倍，合法 OOXML 范围大于 0 且不超过 132。point 规范化到 0.01pt，factor 规范化到 0.00001。

## 方案选择

考虑过三种方案：

1. 直接在 `AddTextOptions` 和 `RichTextParagraph` 上复制 `lineSpacing`、`lineSpacingMultiple`、`paraSpaceBefore`、`paraSpaceAfter`。迁移最直接，但 exact 与 multiple 可同时出现；PptxGenJS 只能靠分支顺序让 exact 获胜，还会接受负行距并写出非法 OOXML。
2. 使用 `spacing` 对象和 `ParagraphLineSpacing` 判别联合。它保留全部公开功能，阻止互斥配置，支持稳定 getter snapshot，并为 paragraph 覆盖/清除提供明确语义。因此采用。
3. 直接暴露 `lnSpc/spcBef/spcAft` 与 `spcPts/spcPct`。它能覆盖更多 OOXML，但把存储单位、schema choice 和 child 顺序泄漏给调用方，不适合作为当前公共值模型。

本库追求功能对等而不是复制 PptxGenJS 的缺陷。adapter 会读取 PptxGenJS 实际生成的 spacing；本库原生 API 对冲突、负数、非有限值和超范围值在 mutation 前失败。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `lineSpacing: 28` 写 `<a:lnSpc><a:spcPts val="2800"/></a:lnSpc>`。
- `lineSpacingMultiple: 1.5` 写 `<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>`。
- `paraSpaceBefore: 6.25` 和 `paraSpaceAfter: 8.5` 分别写 `625` 与 `850`。
- exact 与 multiple 同时存在时只写 exact。
- 四个值为 0 时全部省略。
- 负 `lineSpacing` 会写负 `spcPts`；本库拒绝该无效输出。
- TextProps run 级 spacing 与同段其他 run 冲突时，4.0.1 可能在一个 `a:p` 中写多个 `a:pPr`。本库的 paragraph 值模型只写一个合法 direct `pPr`，不复制该结构缺陷。

## OOXML 映射

```xml
<a:pPr>
  <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="600"/></a:spcBef>
  <a:spcAft><a:spcPts val="800"/></a:spcAft>
  <a:buNone/>
</a:pPr>
```

point 使用 `Math.round(value * 100)` 写入 `spcPts@val`；multiple 使用 `Math.round(factor * 100000)` 写入 `spcPct@val`。exact 与 multiple 是同一个 `lnSpc` choice，公共联合保证最多写一种。

getter 只读取 direct `pPr` 的 direct `lnSpc`、`spcBef` 和 `spcAft`。属性必须是严格十进制整数并落在 schema 范围内。`lnSpc/spcPts` 还原 exact，`lnSpc/spcPct` 还原 multiple；段前/段后只把 `spcPts` 映射到公共 API。零值规范化为无字段，畸形值和未建模的段前/段后 `spcPct` 不伪造成错误语义。

## Lossless patch 与顺序

rich text 替换继续按 paragraph 索引选择 `pPr` 模板。实现只移除 direct `lnSpc`、`spcBef`、`spcAft`，按 DrawingML schema 顺序重建目标 spacing block，并把它插入 `pPr` 的所有 element child 之前。这样 spacing 始终位于 bullet、tab、`defRPr` 和 `extLst` 之前，同时保留 alignment、margin/indent、bullet、tab、默认 run properties、未知属性和未知 child 的相对顺序与字节内容。

显式 `richText` setter 以公共 spacing 值为替换目标，因此未建模的段前/段后百分比不属于保留目标；只读、shape transform、slide 操作等其他 mutation 仍原样保留这些 XML。自闭合 `pPr` 必须安全展开，所有修改继续使用 `LosslessXmlDocument` source-span patch，不以正则重写 XML。

## 验证与错误处理

`ParagraphSpacing` 必须是非数组对象并至少提供 `before`、`after`、`line` 之一，未知字段失败。point 必须有限且位于 0–1584，exact 必须大于 0；factor 必须有限且位于大于 0 到 132。line 对象必须有唯一受支持的 `kind` 和对应数值，未知或交叉字段失败。

`addText()`、`addRichText()` 和 `richText` setter 继续在 transaction 中完成 validation、merge、render 和写回。任何 validation 或 render 失败都必须恢复 package bytes、mutation journal、live shape 内容和 spacing snapshot。

## 测试与验收

1. `addText()` 把默认 spacing 应用到每个 CR/LF paragraph，point 与 multiple 单位正确。
2. `addRichText()` 支持文本框默认、逐字段 paragraph override、单字段清除、整组 `false`、空 paragraph，并与 alignment 和 bullet 组合。
3. getter 正确读取 exact/multiple/before/after，规范化零值，忽略畸形值和未建模的段前/段后百分比，只读不产生 mutation。
4. `shape.richText` 可增加、更换或清除 spacing，同时保留 alignment、bullet、margin/indent、tab、`defRPr`、`extLst`、未知属性/child，且 spacing child 顺序合法。
5. 非法结构、未知字段、NaN/Infinity、负值、互斥 line 字段和越界值在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
6. PptxGenJS 4.0.1 conformance 覆盖 exact、multiple、before/after、小数转换、exact 优先和零值省略；adapter 按实际 OOXML 读取。
7. write/reopen、duplicate、snapshot isolation、全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出一页对比 automatic、exact、multiple、段前和段后间距的 PDF；最终页面无截断或溢出，兼容矩阵把 spacing 标记为已支持。
