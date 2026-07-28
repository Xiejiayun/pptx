# 段落项目符号与编号设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为普通文本和结构化 rich text 增加项目符号与自动编号，使原生创建和已有 PPT 编辑覆盖 PptxGenJS 4.0.1 的标准 bullet、自定义字符、16 种公开编号样式、起始值和 bullet indent。能力必须按 paragraph 表达，并与已经支持的 alignment、rich runs、空 paragraph 和 soft break 正确组合。

本小项不包括多级列表 `indentLevel`、普通段落 margin/first-line indent、tab stop、bullet 颜色/字体/大小、图片 bullet 或 OOXML 超出 PptxGenJS 的编号样式。这些继续作为独立小项；读取和非 rich-text mutation 不得破坏其原始 XML。

## API

```ts
type NumberingStyle =
  | 'alphaLcParenBoth'
  | 'alphaLcParenR'
  | 'alphaLcPeriod'
  | 'alphaUcParenBoth'
  | 'alphaUcParenR'
  | 'alphaUcPeriod'
  | 'arabicParenBoth'
  | 'arabicParenR'
  | 'arabicPeriod'
  | 'arabicPlain'
  | 'romanLcParenBoth'
  | 'romanLcParenR'
  | 'romanLcPeriod'
  | 'romanUcParenBoth'
  | 'romanUcParenR'
  | 'romanUcPeriod';

type ParagraphBullet =
  | boolean
  | {
      readonly kind: 'bullet';
      readonly character?: string;
      readonly indent?: number; // points
    }
  | {
      readonly kind: 'number';
      readonly style?: NumberingStyle;
      readonly startAt?: number;
      readonly indent?: number; // points
    };

interface AddTextOptions extends Partial<Transform> {
  readonly bullet?: ParagraphBullet;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly bullet?: ParagraphBullet;
}

slide.addText('First\nSecond', { bullet: true });

const shape = slide.addRichText(
  [
    { runs: [{ text: 'Default bullet' }] },
    { runs: [{ text: 'Custom symbol' }], bullet: { kind: 'bullet', character: '▶', indent: 18 } },
    {
      runs: [{ text: 'Roman number' }],
      bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
    },
    { runs: [{ text: 'No bullet' }], bullet: false },
  ],
  { bullet: true },
);
```

`true` 是标准 `•` bullet 的创建简写，`false` 明确抑制文本框级默认值。paragraph 未提供 `bullet` 时使用 `AddTextOptions.bullet`；paragraph 值优先。getter 始终返回规范化对象而不是 `true`，没有本地 bullet 时省略该字段。

`character` 是实际 Unicode 字符而不是 PptxGenJS 的十六进制字符串，必须恰好包含一个 Unicode scalar value。这样避免把编码表示泄漏进值模型，也支持 XML 正确转义。默认字符是 `•`。

`indent` 单位为 point，表示 bullet/number 到 paragraph 左边界使用的 hanging margin；默认 27pt，合法范围 0–4032pt。编号默认 `arabicPeriod`、从 1 开始；`startAt` 必须是 1–32767 的整数。

## 方案选择

考虑过三种方案：

1. 直接复制 PptxGenJS 的 `boolean | { type, characterCode, numberType, numberStartAt, indent }`。迁移表面最短，但 4.0.1 的 `type:'bullet'` 会使自定义字符失效，公开 `numberType` 也被运行时代码忽略；复制会固化已确认的缺陷。
2. 使用 `kind` 判别联合和真实 Unicode 字符，并保留 `true/false` 简写。它能表达相同功能，避免互斥字段组合，正确实现公开编号样式且便于 getter 生成稳定 snapshot。因此采用。
3. 暴露 `buChar`、`buAutoNum`、`marL` 等原始 OOXML。覆盖面大，但会泄漏存储 token、破坏单位一致性，并让 XML 顺序和合法性落到调用方。

该 API 追求功能对等而不是复刻 PptxGenJS 的运行时缺陷。adapter 继续通过公开生成结果导入：PptxGenJS 实际写出的 `buChar`/`buAutoNum` 会被正确读取；本库原生创建则按其文档意图正确支持 `romanUcPeriod` 等样式。

## OOXML 映射

标准或自定义 bullet 写入：

```xml
<a:pPr marL="342900" indent="-342900">
  <a:buSzPct val="100000"/>
  <a:buChar char="•"/>
</a:pPr>
```

自动编号写入：

```xml
<a:pPr marL="279400" indent="-279400">
  <a:buSzPct val="100000"/>
  <a:buFont typeface="+mj-lt"/>
  <a:buAutoNum type="romanUcPeriod" startAt="3"/>
</a:pPr>
```

point 使用 `Math.round(value * 12700)` 转为 EMU。active bullet 写 `marL=indent`、`indent=-indent`。没有 bullet 时写 `a:buNone`；清除既有 bullet 时，如果原 `marL/indent` 是 bullet 产生的等值正负 hanging pair，则同时归零，否则保留未来 paragraph-indent 小项尚未建模的自定义属性。

getter 从 direct `pPr` 读取 `buChar` 或 `buAutoNum`。`marL` 为非负有限值时还原 point indent；支持的 `type` 和合法 `startAt` 返回到 snapshot。`buNone` 或没有 bullet choice 时省略 `bullet`。图片 bullet、未知编号 token 和 bullet 的自定义颜色/字体/大小不伪造成错误语义。

## Lossless patch 与顺序

rich text 替换继续按 paragraph 索引选择 `pPr` 模板。实现只移除 bullet choice 相关的 direct children：`buClr*`、`buSz*`、`buFont*`、`buNone`、`buChar`、`buAutoNum` 和 `buBlip`，然后按 DrawingML schema 顺序在 `tabLst/defRPr/extLst` 之前插入新的 block。alignment、spacing、未知属性和未知 child 保持原顺序与字节内容。

active bullet 只更新 `marL/indent` 两个属性；清除时仅对可确认的 bullet hanging pair 归零。自闭合 `pPr` 必须安全展开。实现使用 `LosslessXmlDocument` source-span patch，不以正则改写 XML。

显式 `richText` setter 以公开 bullet 值为替换目标，因此未建模的图片 bullet 或 bullet 颜色/字体/大小不属于保留目标；只读、shape transform、slide 操作等其他 mutation 仍原样保留这些 XML。

## 验证与错误处理

对象必须有唯一受支持的 `kind`，未知字段明确失败。`character` 的空字符串、多 scalar、控制字符和非法 XML 字符失败；`indent` 必须有限且在 0–4032pt；`startAt` 必须为 1–32767 的整数；`style` 必须属于公开 16 种 token。所有失败发生在 package mutation 前。

`addText()`、`addRichText()` 和 `richText` setter 继续在 transaction 中完成验证、render 和写回。内外层 rollback 必须恢复 bytes、mutation journal、live shape 内容和 bullet snapshot。

## 测试与验收

1. `addText()` 将 `bullet:true` 应用到每个 CR/LF paragraph，默认 `•`、27pt hanging indent 正确。
2. `addRichText()` 支持文本框默认、逐段 custom Unicode bullet、16 种编号 style、startAt、indent 和 `false` 覆盖；空 paragraph 也能带 bullet。
3. getter 正确读取 `buChar`、`buAutoNum`、indent 和编号字段；`buNone`、未知编号和图片 bullet 不伪造值，只读不产生 mutation。
4. `shape.richText` 可增加、更换或清除 bullet，同时保留 alignment、spacing、tab、未知属性/child 和非 bullet indent；schema child 顺序合法。
5. 非法结构、字符、style、startAt、indent 在 mutation 前失败；内外层 rollback 不改变 bytes/journal/identity。
6. PptxGenJS 4.0.1 真实输出覆盖标准 bullet、自定义 `characterCode`、公开 `numberType` 的已知回退和旧 `style` 的实际编号输出；adapter 按生成的 OOXML 读取。
7. write/reopen、duplicate、全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出包含标准/custom bullet、字母/阿拉伯/罗马编号的一页 PDF；兼容矩阵把 bullet/numbering 标记为已支持，并记录多级列表仍待实现。
