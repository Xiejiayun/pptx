# PPTX OOXML

一个面向现有 PPTX 的 TypeScript 双向编辑内核。它把演示文稿作为 OPC package graph 读取，通过 source-span XML patch 做局部修改，并默认保留未知 OOXML、扩展节点和二进制部件。

当前实施进度和功能截图见 [docs/implementation-progress.md](./docs/implementation-progress.md)。完整路线图见 [PLAN.md](./PLAN.md)。

```sh
npm install @jiayunxie/pptx@next
```

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = await PptxDocument.open('input.pptx');
document.slides[0].title.text = 'Updated';
await document.writeFile('output.pptx');
```

## 创建和编辑形状填充、线条、箭头与链接

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const shape = slide.addShape('roundRect', {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 20,
  },
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '1F4E78' },
    transparency: 10,
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  hyperlink: {
    url: 'https://example.com/docs',
    tooltip: '打开文档',
  },
});

shape.fill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
shape.fill = { kind: 'none' };
shape.fill = undefined;
shape.line = { kind: 'line', color: { kind: 'scheme', value: 'accent2' } };
shape.line = { kind: 'none' };
shape.line = undefined;
shape.arrows = { begin: 'diamond' }; // 同时清除省略的 end
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined; // 清除两端，保留线条样式
document.addSlide(); // 创建第 2 页作为内部链接目标
shape.hyperlink = { slide: 2, tooltip: '前往详情' };
shape.hyperlink = { url: 'mailto:team@example.com', tooltip: '' };
shape.hyperlink = undefined;
```

`ShapeModel.fill` 支持 direct solid/no-fill 的创建、读取、编辑与清除。`{ kind: 'none' }` 写入明确的 direct no-fill，`undefined` 只清除 direct fill state；gradient、pattern、picture 和 group fill 不属于这个 simple-fill API。

`ShapeModel.line` 支持 direct none/solid line 的创建、读取、编辑与清除，包括 sRGB/theme color、0–100% transparency、0–1584pt width 和 8 种 preset dash。省略 width/dash 默认 1pt/solid；`undefined` 只清除 line 的 width/fill/dash，同时保留 line 容器、箭头、join 和扩展节点。

`AddShapeOptions.arrows` 与 `ShapeModel.arrows` 支持 begin/end 的 `none | arrow | diamond | oval | stealth | triangle`。快照与输入脱离；赋值采用 whole replacement，缺失的一端会被清除，显式 `none` 则保留对应 direct endpoint。`undefined` 只清除两端而保留 line，反向的 `shape.line = undefined` 也保留 arrows。只创建 arrows 不会隐式生成颜色、宽度或 dash；已有合法 `w` / `len` size 会在类型编辑中无损保留，但 size 创建/读取/编辑尚未公开。

`AddShapeOptions.hyperlink` 与 `ShapeModel.hyperlink` 支持整个 preset shape 的 click URL 或内部页链接。输入必须恰好包含一个非空 `url` 或一个当前文稿内的一基 `slide`；`tooltip` 可省略，也可显式为空。Getter 返回 detached frozen snapshot，setter 采用 whole replacement，同值赋值为 exact no-op，`undefined` 清除 click link。内部关系按目标页 identity 保存，移动或在目标前插删页面只更新 getter ordinal；复制 self-link 会指向副本自身，删除目标页会清理相关 click/hover，shared relationship 则按引用 clone-on-write 与回收。外部链接产生 validator 的预期可移植性 warning。Hover 编辑、text-run/table/image/chart/media 链接创建、action navigation、shadow、adjustment、custom geometry、advanced line fill/custom dash 和 percentage positions 仍待后续小项。

## 开发

```sh
pnpm install
pnpm check
pnpm build
```

Node.js 20+ 或现代浏览器；TypeScript strict mode。浏览器支持 `Blob`、`File`、Web `ReadableStream` 输入，以及 `writeBlob()` 和 `download()` 输出。

## Workspace packages

npm 用户只需安装 `@jiayunxie/pptx`。以下是仓库内部模块边界，不需要分别安装：

- `@pptx/sdk`：统一的打开、编辑、验证和保存 API。
- `@pptx/opc`：ZIP、content types、relationships 和 package graph。
- `@pptx/lossless-xml`：source-span XML tree 与最小 patch。
- `@pptx/model`：slide 和常规对象语义模型。
- `@pptx/codecs`：Master/Layout/Theme、Gradient/Transparency、Media。
- `@pptx/pptxgenjs-adapter`：PptxGenJS 公开输出适配。
- `@pptx/testkit`：part hash diff、mutation isolation 和 LibreOffice helper。
- `@pptx/cli`：`pptx-inspect` 离线 inspection CLI。
- `@pptx/plugin-transitions`：转场、自动换页与声音。
- `@pptx/plugin-animations`：动画/媒体 timing tree。
- `@pptx/plugin-advanced-charts`：组合/现代图表及高级 series 功能。
- `@pptx/plugin-smartart`：SmartArt part set 与节点编辑。

## CLI

```sh
npx @jiayunxie/pptx@next --json doctor
pptx-inspect --json package validate deck.pptx
pptx-inspect --json slides list deck.pptx
```

CLI 写操作要求明确输出路径，并支持 `--dry-run`。完整命令见 [packages/cli/README.md](./packages/cli/README.md)。
