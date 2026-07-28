# Rich Text 语言设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 plain text 和结构化 rich text 增加 PptxGenJS 4.0.1 `lang` 对等能力：创建文本框时可声明默认语言，单个 rich-text run 可继承或覆盖该语言，`ShapeModel.richText` 可读取和整体替换 run 的 direct 语言。语言与现有字体、字号、粗斜体、颜色、baseline、character spacing、glow、highlight、outline、underline、strike、soft break 和 paragraph options 组合。

本小项不公开 `altLang`、RTL、East Asian/complex-script 字体、proofing 状态或 paragraph/default-text-style 继承解析。读取和非 rich-text mutation 不得破坏这些原始 XML。

## API

```ts
interface AddTextOptions {
  readonly lang?: string;
}

interface RichTextRunStyle {
  readonly lang?: string;
}

slide.addText('Bonjour', { lang: 'fr-CA' });

slide.addRichText(
  [{
    runs: [
      { text: 'Français' },
      { text: ' Deutsch', style: { lang: 'de-DE' } },
    ],
  }],
  { lang: 'fr-CA' },
);
```

`AddTextOptions.lang` 是 plain/rich 创建时的文本框级默认语言。rich run 未提供 `style.lang` 时继承该默认值，run 值优先。`ShapeModel.richText` setter 没有文本框默认参数：run 省略 `lang` 时使用现有 `en-US` 生成默认；提供 `lang` 时写 direct 语言。

语言值是不透明的非空 string，不裁剪、不改大小写，也不尝试用不完整的正则重写 BCP 47。写入会拒绝 XML 1.0 非法控制字符并转义 attribute metacharacters，因而支持合法的扩展/private-use tag，同时不会重现 PptxGenJS 的未转义 XML 缺陷。

## 方案选择

考虑过三种方案：

1. `AddTextOptions.lang` 加 `RichTextRunStyle.lang`；采用此方案。它完整表达 PptxGenJS outer default 与 run override，也让打开后的 direct run 语言可编辑。
2. 只增加 `RichTextRunStyle.lang`。实现最小，但 plain text 无法声明语言，rich runs 也必须重复默认值，未达到 PptxGenJS 对等。
3. 引入 `{ primary, alternate }` 语言对象。可公开 `altLang`，但 PptxGenJS 的公共 surface 只有 `lang`，而 alternate language 的继承与编辑边界需要独立设计，超出本小项。

公共名称保留 `lang`，避免迁移时无意义改名；内部 normalized 字段使用 `language` 以保持代码可读性。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- 未指定 outer/run `lang` 时，run 写 `lang="en-US"`，不写 `altLang`；`endParaRPr` 写 `lang="en-US"`。
- outer `lang: 'fr-CA'` 会传播到所有未显式设置语言的 runs；这些 runs 写 `lang="fr-CA" altLang="en-US"`。
- run `lang: 'de-DE'` 覆盖 outer 值并写 `lang="de-DE" altLang="en-US"`。
- 空字符串、0 或 false 是 falsy：有 outer 值时继承 outer，没有 outer 值时回退 `en-US`。
- `endParaRPr` 使用 outer `lang`，不跟随单个 run override，也不写 `altLang`。
- PptxGenJS 会把 truthy 字符串未经 XML attribute escaping 原样写入；例如 `fr-CA&x` 会生成无法由 XML parser 读取的 slide part。

本库匹配有效 outer/run/default/end-paragraph 输出，但严格拒绝非 string 和空 string，并转义有效 string 中的 XML metacharacters。显式 `lang: 'en-US'` 仍视为显式值并带 `altLang="en-US"`，与 PptxGenJS 一致。

## OOXML 映射

省略语言：

```xml
<a:rPr lang="en-US" dirty="0">...</a:rPr>
<a:endParaRPr lang="en-US" dirty="0"/>
```

outer 默认和 run override：

```xml
<a:rPr lang="fr-CA" altLang="en-US" dirty="0">...</a:rPr>
<a:rPr lang="de-DE" altLang="en-US" dirty="0">...</a:rPr>
<a:endParaRPr lang="fr-CA" dirty="0"/>
```

plain `addText()` 把 normalized outer language 传给每个非空 run，并传给每个空/非空 paragraph 的 `endParaRPr`。`addRichText()` 把它作为 `renderRichTextParagraphs()` 的 default language；`renderRun()` 按 `style.lang ?? defaultLanguage ?? 'en-US'` 解析值，并仅在前两者存在时写 `altLang="en-US"`。

getter 只读取 direct `rPr@lang`。非空值原样暴露到 `RichTextRunStyle.lang`；属性缺失或空值不返回字段。`altLang` 不影响 getter。读取不解析 master/layout/list/default-run 继承，也不因为默认语言是 `en-US` 而隐藏 direct 属性。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 `lang`/`altLang`。plain `.text` 继续复制第一 paragraph/run 模板，因此原语言属性随模板保留。

`shape.richText` setter 整体替换 runs：新 run 提供 `lang` 时写该 direct 语言与 `altLang="en-US"`，未提供时写默认 `lang="en-US"` 且不写 `altLang`。既有 run 的未公开 `altLang` 不逐属性保留；paragraph `pPr`、第一段 `endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。若模板没有 `endParaRPr`，replacement 使用 `en-US` 默认。

创建时，outer language 决定所有 paragraph 的 `endParaRPr`。run override 不改变 `endParaRPr`。打开已有文件后整体替换 rich text 时，现有第一段 `endParaRPr` 仍按当前 lossless 模板规则复制到所有 replacement paragraphs。

## 验证与错误处理

outer/run language 必须是长度大于 0 的 string，并且不能包含 XML 1.0 禁止的控制字符。number、boolean、null、object、array、空 string 或非法 XML 字符明确失败；空白和大小写不被静默规范化。

所有创建参数与所有 runs 在 package mutation 前完成验证。单个 run 或 outer language 失败时，不得留下新 shape、part bytes、mutation journal、live shape snapshot 或对象 identity 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. plain `addText()` 覆盖默认 `en-US`、outer language、多 paragraph、空 paragraph、attribute escaping 和 `endParaRPr`。
2. `addRichText()` 覆盖 outer inheritance、run override、显式 `en-US`、无 outer 默认、空 runs，以及与现有 typography/paragraph options 组合。
3. getter 读取 direct `lang`，缺失/空 attribute 不伪造字段，保留并忽略 `altLang`，只读不产生 mutation。
4. `shape.richText` 增加、更换和清除 run language；plain `.text` 保留模板语言；write/reopen、duplicate 与 transaction rollback 一致。
5. 非 string、空 string 和非法 XML 字符在 mutation 前失败，错误上下文区分 outer 与具体 paragraph/run。
6. PptxGenJS 4.0.1 conformance 覆盖 default、outer inheritance、run override、falsy fallback、显式 `en-US`、`endParaRPr` 与真实输出导入。
7. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出多语言对照页；原生文件与同版本 PptxGenJS 文件的有效语言 OOXML 一致。
