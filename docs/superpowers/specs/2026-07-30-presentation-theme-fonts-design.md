# Presentation Theme Fonts Design

日期：2026-07-30
状态：已批准实施（用户已授权自主选择最佳方案并持续推进）

## 目标与范围

补齐 PptxGenJS 4.0.1 顶层 `theme.headFontFace` 与 `theme.bodyFontFace` 的原生能力：从零创建时可指定标题和正文字体，打开已有 PPTX 后可读取、整体替换或局部编辑 presentation 直接引用 theme 的 Latin major/minor 字体，并在保存、重新打开和 adapter 导入后保持。

本小项只拥有主题 font scheme 中 `majorFont/latin@typeface` 与 `minorFont/latin@typeface`。它不增加主题颜色、East Asian/complex-script/script-specific font 编辑，不定义或选择多个命名主题，不改变 master/layout/notes 的关系拓扑，也不把主题字体实体化到每个 run。现有 `ThemeModel.colors`、`setColor()`、master/layout/theme 创建复制删除关系和所有未知 theme 内容继续独立保留。

## 公共 API

```ts
export interface PresentationTheme {
  readonly headFontFace: string;
  readonly bodyFontFace: string;
}

export interface PresentationThemeOptions {
  readonly headFontFace?: string;
  readonly bodyFontFace?: string;
}

export interface ThemeFontSnapshot {
  readonly majorLatin: string;
  readonly minorLatin: string;
}

export interface ThemeFontUpdate {
  readonly majorLatin?: string;
  readonly minorLatin?: string;
}

interface CreatePresentationOptions {
  readonly theme?: PresentationThemeOptions;
}

const document = PptxDocument.create({
  theme: {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
  },
});

document.theme = {
  headFontFace: 'Noto Sans Display',
  bodyFontFace: 'Noto Sans',
};

document.themes[0]?.setFonts({ minorLatin: 'Noto Sans' });
```

`document.theme` getter 返回 detached 的完整 `{ headFontFace, bodyFontFace }` 快照；若 presentation 没有唯一、安全、可解析的直接 theme，则返回 `undefined`。setter 是 PptxGenJS 风格的整组替换：缺少 `headFontFace` 时使用 `Calibri Light`，缺少 `bodyFontFace` 时使用 `Calibri`。`ThemeModel.setFonts()` 是面向已有文件的局部 direct editor，省略的一侧保持不变。

零参数 `PptxDocument.create()` 继续使用当前 canonical `Aptos Display` / `Aptos`，不为了复制 PptxGenJS 品牌时代的默认值而改变已有输出。只有显式提供 `theme` 对象或赋值 `document.theme` 时才应用 PptxGenJS 4.0.1 的 partial-field 回退。显式空对象因此写入 `Calibri Light` / `Calibri`；runtime `theme: undefined` 等同省略并保持 canonical Aptos。

## 输入契约

字体名必须是包含至少一个非空白字符的 string，并且不能包含 XML 1.0 非法字符。合法值保持原字符串，不 trim、不改大小写，XML 元字符正确转义。空字符串、纯空白、非 string、数组、accessor 属性、symbol 属性、原型污染对象和未知 own key 在任何 package mutation 前拒绝。

`PresentationThemeOptions` 只允许 `headFontFace`、`bodyFontFace`；`ThemeFontUpdate` 只允许 `majorLatin`、`minorLatin`，且必须至少提供一个非 `undefined` own data property。输入在调用时 descriptor-safe 读取并立即脱离 caller，后续修改原对象不影响文稿。

PptxGenJS 4.0.1 使用 truthy 判断：空字符串回退默认，纯空白会原样写入，非 string truthy 值可被字符串插值，XML 元字符也不会 escape。原生 API 只对合法、明确的字体名提供功能对等，不复制这些会产生无意义或 malformed OOXML 的运行时缺陷。

## 方案选择

考虑三种实现：

1. 只在 `create.ts` 的 `THEME_XML` 模板中做字符串替换。实现短，但无法编辑已有文件，容易破坏 escaping，也绕开已经存在的 theme codec ownership。
2. 在 SDK 中直接选择 `document.themes[0]`。它可复用当前 model，却把“第一个 theme part”误当 presentation theme；多 master、theme override、孤立 theme part 或非标准 part 顺序都会选错。
3. 扩展 `MasterLayoutThemeCodec`：严格解析 presentation 的直接 theme relationship，在 `ThemeModel` 中提供 font snapshot/patch，再由 SDK 暴露 PptxGenJS 风格整组属性。它同时满足创建、编辑、关系准确性和未知内容保留，因此采用。

## OOXML ownership 与定位

顶层 `document.theme` 只跟随 presentation part 上唯一 direct relationship：

```text
presentation.xml
  -- officeDocument/2006/relationships/theme --> theme part
```

relationship 必须是 internal、resolved target 存在且目标 content type 为 `application/vnd.openxmlformats-officedocument.theme+xml`。零条、重复、external、dangling 或 wrong-content-type 状态不回退到 master、layout、notes master 或“唯一 theme 文件”；getter 返回 `undefined`，mutation 明确失败且保持零变化。

目标 part 必须有唯一 DrawingML namespace 的 direct 结构：

```text
a:theme
  / a:themeElements
    / a:fontScheme
      / a:majorFont / a:latin @typeface
      / a:minorFont / a:latin @typeface
```

查找按 expanded namespace 与 direct-child 层级，不按固定 `a` prefix、固定 part URI 或任意 descendant 猜测。alternate prefix 和局部 namespace declaration 合法；wrong namespace、重复 direct 节点、descendant impostor、重复 unqualified `typeface` attribute或 element child 都不被误读。Malformed XML 继续由 lossless parser 明确抛错，不被吞成“缺少字体”。

setter 只替换唯一 direct Latin element 的 unqualified `typeface` 值；单纯缺少该 attribute 时可安全补上。wrong-namespace 同名 attribute 保留。major/minor 容器或 Latin element 缺失、重复或 namespace 错误时不创建大段 font scheme，而是拒绝 mutation，避免猜测 child order 与主题继承。未知 attributes、`panose`、`pitchFamily`、`charset`、`ea`、`cs`、script-specific `font`、color scheme、format scheme、object defaults、extensions、comments、whitespace 和其他 parts 均保持原字节。

## 数据流与原子性

创建流程先解析 explicit `theme` 输入，再建立 canonical package，最后在同一 document transaction 内通过 presentation theme relationship 修改 theme1。失败不会返回半初始化文稿。metadata、slide size 和 RTL 的现有创建顺序不受影响。

`document.theme` getter 通过 `MasterLayoutThemeCodec.presentationTheme` 获取稳定 `ThemeModel`，再把 strict `fonts` 快照映射为 PptxGenJS 名称。setter 将 partial input 归一化为完整 pair 后调用 `ThemeModel.setFonts()`。`ThemeModel.setFonts()` 在一个 OPC transaction 内验证完整目标结构、比较当前值并进行最小 source-span attribute patch；两值相同和 partial same-value 都是 exact bytes/journal no-op。外层 transaction 失败会回滚两个字体修改。

同一 theme part 被 presentation、master 和 notes master 共享时，编辑该 part 会按 OOXML 引用语义影响所有引用者；本小项不 clone-on-write，因为顶层 presentation theme 本来就是全局意图。调用方若需要编辑非主 theme，可继续从 `document.themes` 选择具体 `ThemeModel`。

## PptxGenJS 4.0.1 基线

真实公开输出证明：

- 未设置 `theme` 或设置 `{}`：major Latin 为 `Calibri Light`，minor Latin 为 `Calibri`；
- 只提供 `headFontFace`：major 使用 custom，minor 回退 `Calibri`；
- 只提供 `bodyFontFace`：major 回退 `Calibri Light`，minor 使用 custom；
- 两者都提供：两个 `a:latin@typeface` 都使用 custom；
- custom Latin 不写 `panose`，fallback Calibri 带 PptxGenJS 的固定 `panose`；
- assignment 保留 caller 对象引用并在 `write()` 时读取，原生实现则立即 detached；
- adapter 只调用公开 `write()`，导入最终 theme XML，不读取 `_theme`。

原生对等以最终 major/minor Latin typeface 语义为准，不要求复制 PptxGenJS 整份 Office Theme 模板、script font 列表、panose 差异或未转义缺陷。PptxGenJS default 输出导入后 `document.theme` 必须为 `Calibri Light` / `Calibri`；custom 与 partial 输出也必须精确读取并在 native write/reopen 后保持。

## 测试与验收

1. Codec fixture 覆盖 canonical、alternate URI/prefix、局部 namespace、额外 theme parts、presentation direct relationship、master-only relationship和 detached theme part；只选择唯一 direct presentation theme。
2. `ThemeModel.fonts` 覆盖 valid、missing/repeated/wrong-namespace chain、missing/duplicate typeface和 descendant impostor，并返回 detached snapshot 或 `undefined`；malformed XML 明确抛出 parse error。
3. `setFonts()` 覆盖 major-only、minor-only、both、missing-attribute repair、XML escaping、same-value no-op、empty-update rejection、invalid zero mutation、unknown-state preservation和外层 rollback。
4. `PptxDocument.create()` 覆盖 omitted、runtime `undefined`、`{}`、head-only、body-only、both、冻结/null-prototype input、caller 后续 mutation及全部非法输入；默认 canonical bytes保持不变。
5. `document.theme` 覆盖读取、整组替换、partial fallback、write/reopen、stable `ThemeModel` identity，以及缺失/重复/external/dangling/wrong-content-type relationship 的安全行为。
6. PptxGenJS 4.0.1 conformance 仅用 public constructor、`theme` setter、`addSlide()`、`write()`，覆盖 default、empty、head-only、body-only、both，并通过 adapter 验证 final typeface。
7. npm tarball 的 Node、browser 和 declaration smoke 覆盖 `CreatePresentationOptions.theme`、live getter/setter、`ThemeModel.setFonts()` 和 CLI 版本。
8. `pptx-inspect` 验证 native/PptxGenJS 文件为 PowerPoint 2010 profile 0 error / 0 warning；part diff 证明只改目标 theme part，同值与 reopen 为零变化。
9. LibreOffice headless 无修复打开并导出 native custom、edited、reopened 和 PptxGenJS custom 文件；固定可用字体下的渲染和页面尺寸一致，Presentations render/overflow 检查通过。
10. 全量 typecheck、test、performance、browser build、真实 tarball smoke、git diff check 和远端 `origin/main...HEAD = 0 0` 全部通过。
