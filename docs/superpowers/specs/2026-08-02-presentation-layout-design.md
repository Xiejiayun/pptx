# Presentation `presLayout` Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 document instance 的只读 `presLayout` 能力，让 native document 可以用一个稳定、只读、可序列化的对象读取当前幻灯片画布：

```ts
interface PresentationLayout {
  readonly name: PresentationLayoutName;
  readonly width: Emu;
  readonly height: Emu;
}
```

- `PptxDocument.presLayout` 对 create/open/edit/write/reopen 后的文稿都可用；
- 默认、`4:3`、`16:9`、`16:10`、`wide` 与任意 OOXML 合法自定义尺寸都可读取；
- width/height 与现有 `slideSize` 一样使用 EMU，避免同一 document 出现两套单位；
- getter 每次从 `slideSize` 计算并返回 detached snapshot，不新增 registry、cache 或第二份可变尺寸状态；
- 修改 `slideSize` 后下一次读取立即反映新尺寸；
- 读取不产生 OPC mutation，写出前后与 reopen 后尺寸保持一致。

本项不新增 PptxGenJS 的 `layout` setter 或 `defineLayout()` registry。Native 已分别通过 `create({ slideSize })` 和 `document.slideSize` 覆盖其最终文件能力；自定义 layout 名称不写入 OOXML，打开既有 PPTX 时无法可靠恢复，因此不伪造持久名称。本项也不处理其他 runtime constants、output type/stream/compression、advanced text/table 或 `tableToSlides`。

## 2. 当前状态与 PptxGenJS 权威行为

Native 已有：

- `BuiltInSlideSize = '4:3' | '16:9' | '16:10' | 'wide'`；
- `create({ slideSize })` 接受四种内建名称或 `{ width, height }`；
- `document.slideSize` getter/setter 读取和编辑 `p:sldSz`；
- `SlideSize` 使用只读 EMU 宽高并返回 detached snapshot；
- 尺寸验证覆盖 1–56 inch、transaction rollback、六格式与 write/reopen。

PptxGenJS 4.0.1 declaration 公开 `readonly presLayout: PresLayout`，其中 `PresLayout` 的公开字段只有 `name`、`width`、`height`。Public runtime probe 得到：

| layout | `presLayout.name` | width × height |
| --- | --- | --- |
| 默认 / `LAYOUT_16x9` | `screen16x9` | 9,144,000 × 5,143,500 |
| `LAYOUT_4x3` | `screen4x3` | 9,144,000 × 6,858,000 |
| `LAYOUT_16x10` | `screen16x10` | 9,144,000 × 5,715,000 |
| `LAYOUT_WIDE` | `custom` | 12,192,000 × 6,858,000 |
| `defineLayout({ name: 'CUSTOM', width: 11.7, height: 8.3 })` | `CUSTOM` | 10,698,480 × 7,589,520 |

尽管 `defineLayout()` 输入 inch，getter 的运行时 width/height 是 EMU。PptxGenJS getter 没有 setter，但返回同一个内部可变对象；runtime 还可见 `_sizeW` / `_sizeH`，而 declaration 已把这两个字段注释掉，它们不是 typed public surface。Native 不复制内部 alias、私有字段或修改 getter 结果即可破坏 write 的行为。

## 3. 方案比较

### A. 直接把 `slideSize` 改名或返回其对象

代码最少，但缺少 PptxGenJS 的 `name`，也会让两个公开属性的职责不清。拒绝。

### B. 增加可变命名 layout registry

可以保留同一进程内 `defineLayout()` 的任意名称，但名称不会写入 PPTX，open/reopen 后必然丢失；registry 还会与 `slideSize` 形成两个尺寸来源。拒绝。

### C. 从 canonical `slideSize` 派生只读 `presLayout`（采用）

新增一个纯函数把当前宽高投影为 `{ name, width, height }`。三个标准尺寸映射到 `screen4x3`、`screen16x9`、`screen16x10`；PptxGenJS 自己也把 wide 命名为 `custom`，所以 wide 与其他尺寸统一映射为 `custom`。getter 只依赖 OOXML 中真实存在的 `p:sldSz`，create/open/edit/write/reopen 语义一致。

## 4. 公共 API 与命名规则

SDK 新增 browser-safe 模块：

```ts
export type PresentationLayoutName =
  | 'screen4x3'
  | 'screen16x9'
  | 'screen16x10'
  | 'custom';

export interface PresentationLayout {
  readonly name: PresentationLayoutName;
  readonly width: Emu;
  readonly height: Emu;
}

export class PptxDocument extends PresentationModel {
  get presLayout(): PresentationLayout;
}
```

`PresentationLayout` 使用 native descriptive type name，不复制 PptxGenJS namespace 结构；迁移入口仍保留 exact property name `presLayout`。root `@jiayunxie/pptx` 通过现有 SDK export chain 暴露两个类型和 getter。

名称只按宽高精确匹配：

- 9,144,000 × 6,858,000 → `screen4x3`；
- 9,144,000 × 5,143,500 → `screen16x9`；
- 9,144,000 × 5,715,000 → `screen16x10`；
- 其他合法尺寸（包括 12,192,000 × 6,858,000 wide）→ `custom`。

精确匹配避免把近似比例或不同物理尺寸错误标成标准布局。自定义尺寸恰好等于标准尺寸时，文件语义与标准尺寸不可区分，因此返回对应 canonical 名称。

## 5. 状态、错误与 mutation 语义

- getter 调用一次现有 `slideSize`，然后构造新 plain object；重复读取值相等但 identity 不同；
- 返回字段在 TypeScript 中只读；runtime snapshot 不冻结，保持项目现有 detached-value 模式，修改旧 snapshot 不影响 document；
- getter descriptor 没有 setter，runtime strict assignment 与 PptxGenJS 一样失败；
- malformed package 缺少或包含非法 `p:sldSz` 时，沿用 `slideSize` 的现有 `PackageError`，不增加 fallback；
- 读取不修改 package parts、relationships、mutation journal、diagnostics 或 model cache；
- `slideSize` setter transaction 失败后，`presLayout` 继续返回修改前的值；
- getter 不改变 slide、shape、master、layout、notes page 或已有内容的坐标。

## 6. PptxGenJS 对等边界

本项对等的是 typed public `presLayout` 的读取能力、三个字段、EMU 数值、内建布局名称规则、getter-only descriptor 与 write 稳定性。Native 有意改进两处实现细节：

1. 不暴露未声明的 `_sizeW` / `_sizeH`；
2. 返回 detached snapshot，而不是可破坏内部状态的 mutable alias。

PptxGenJS 自定义 registry name 只存在于运行时；Native 对任意非标准尺寸返回 `custom`。这不是文件能力缺口，因为名称不进入 OOXML，最终 PPTX 的 `p:sldSz` 完全对等。文档必须明确此边界，不能宣称 native 支持可恢复的命名 layout registry。

## 7. 验证策略

### Source 与生命周期

- 纯映射函数覆盖四种内建尺寸、自定义尺寸、精确匹配与 detached snapshot；
- document 默认/create/open/edit/write/reopen 返回预期名称与 EMU；
- getter descriptor 无 setter，TypeScript 拒绝 property/field assignment；
- 读取不产生 package mutation，旧 snapshot 的运行时修改不影响下一次读取；
- malformed `p:sldSz` 继续抛出现有错误。

### PptxGenJS public conformance

- 只用 public constructor、`layout`、`defineLayout()`、`presLayout` 和 `write()`；
- 比较默认、四种内建与自定义布局的公开 `name` / width / height；
- 证明双方 getter 都没有 setter并在 write 前后稳定；
- 对 custom name 只比较最终宽高，并显式验证 native canonical name 为 `custom`。

### Packed runtime

- actual tarball Node consumer、browser conditional export 与 declaration consumer 读取 `presLayout`；
- TypeScript consumer 接受 `PresentationLayout` / `PresentationLayoutName`，拒绝 document property 和 snapshot field assignment；
- real Chrome create/edit/writeBlob/reopen 验证相同值且 console/page/network errors 为零；
- installed CLI open/inspect evidence包含画布宽高，不增加专用 CLI 命令。

## 8. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. SDK/root/actual tarball 的 Node、browser 与 declarations 都公开 `presLayout`；
2. 默认、四种内建、自定义、edit、write/reopen、mutation isolation 与 malformed input 全部有永久测试；
3. PptxGenJS public runtime 行为有对照证据，私有字段与 custom-name 边界有明确记录；
4. focused/full Vitest、performance、两套 TypeScript、Node/browser bundle、declaration、pack、CLI 与 real Chrome 门禁全部通过；
5. 六份 release 文档更新，不再把 `presLayout` 列为缺口，但继续保留其余 runtime/output/content 缺口；
6. 每个实施任务独立 review、commit、push，并确认远端同步。
