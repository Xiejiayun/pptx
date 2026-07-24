# PPTX OOXML

一个面向现有 PPTX 的 TypeScript 双向编辑内核。它把演示文稿作为 OPC package graph 读取，通过 source-span XML patch 做局部修改，并默认保留未知 OOXML、扩展节点和二进制部件。

当前实施进度和功能截图见 [docs/implementation-progress.md](./docs/implementation-progress.md)。完整路线图见 [PLAN.md](./PLAN.md)。

```ts
import { PptxDocument } from '@pptx/sdk';

const document = await PptxDocument.open('input.pptx');
document.slides[0].title.text = 'Updated';
await document.writeFile('output.pptx');
```

## 开发

```sh
pnpm install
pnpm check
pnpm build
```

Node.js 20+；TypeScript strict mode。

## Packages

- `@pptx/sdk`：统一的打开、编辑、验证和保存 API。
- `@pptx/opc`：ZIP、content types、relationships 和 package graph。
- `@pptx/lossless-xml`：source-span XML tree 与最小 patch。
- `@pptx/model`：slide 和常规对象语义模型。
- `@pptx/codecs`：Master/Layout/Theme、Gradient/Transparency、Media。
- `@pptx/pptxgenjs-adapter`：PptxGenJS 公开输出适配。
- `@pptx/testkit`：part hash diff、mutation isolation 和 LibreOffice helper。
- `@pptx/cli`：`pptx-inspect` 离线 inspection CLI。

## CLI

```sh
pnpm build
pptx-inspect --json doctor
pptx-inspect --json package validate deck.pptx
pptx-inspect --json slides list deck.pptx
```

CLI 写操作要求明确输出路径，并支持 `--dry-run`。完整命令见 [packages/cli/README.md](./packages/cli/README.md)。
