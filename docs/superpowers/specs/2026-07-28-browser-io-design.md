# 浏览器 I/O 与条件导出设计

日期：2026-07-28  
状态：已批准实施

## 目标

让同一个 `@jiayunxie/pptx` 导入路径在 Node.js 20+ 和现代浏览器中工作。浏览器能够从 `Uint8Array`、`ArrayBuffer`、`Blob`/`File`、Web `ReadableStream` 打开演示文稿，并输出 `Uint8Array`、`Blob` 或触发下载；Node 继续支持文件路径、Node readable stream 和 `writeFile()`。

本小项负责运行时可移植性与发布入口，不承诺 ZIP 流式常量内存，也不增加新的 PowerPoint 语义功能。

## 方案选择

考虑过三种方案：

1. Node/Browser 各维护一个 `PptxDocument` 子类。类型清晰，但 adapter 返回类型和 `instanceof` 行为容易分裂。
2. 为所有文件和加密操作注入环境 adapter。长期最灵活，但当前只有少量环境边界，会提前引入较重配置。
3. 使用一个环境无关 `PptxDocument`，只在实际调用路径文件 I/O 时动态加载 Node 模块；浏览器条件导出使用独立 browser bundle。该方案保持类 identity、现有 Node API 和单一公共入口，因此采用此方案。

## 公共 API

`PptxInput` 扩展为：

- `string`：Node 本地路径；浏览器中明确报错。
- `Uint8Array`、`ArrayBuffer`。
- `Blob`/`File`。
- Web `ReadableStream`。
- `AsyncIterable` 字节流，覆盖 Node readable stream。

`PptxDocument` 新增：

- `writeBlob(options?)`：返回带正确演示格式 MIME type 的 `Blob`。
- `download(fileName?, options?)`：浏览器创建临时 object URL 并触发下载；没有 DOM 时明确报错。

`writeFile()` 和 `openPptxStream()` 保持 Node 能力，但 Node 模块不在顶层静态导入。格式 profile 增加 package 文件 MIME type，供 Blob 和下载复用。

## 可移植媒体与插件

公共 browser bundle 必须包含所有官方能力，因此同时清理会被根入口导出的 Node-only 代码：

- Media 接受 `Blob`、Web stream 和 async iterable；本地路径只在 Node 动态读取。
- Media 哈希去重使用 Web Crypto SHA-256，保持 Node/browser 结果一致。
- 默认 poster 使用字节常量，不依赖 `Buffer`。
- SmartArt UUID 使用标准 `globalThis.crypto.randomUUID()`。

CLI 和 testkit 仍是 Node-only，不进入 browser bundle。

## 构建与发布

- 保留 `dist/index.js` Node ESM 和 `dist/cli.js`。
- 新增 `dist/browser.js`，以 browser platform 构建且不注入 `node:module` banner。
- `package.json#exports["."]` 增加 `browser` 条件，类型入口保持统一。
- browser bundle 不得包含静态 `node:` import。

## 错误处理

- 浏览器传入本地路径或调用 `writeFile()`：抛出说明仅 Node 支持的错误。
- 不支持的 stream chunk：抛出 `TypeError`，不静默字符串化。
- 缺少 Web Crypto：媒体去重操作明确失败；打开、读取和普通编辑不受影响。
- `download()` 缺少 DOM/URL API：明确失败且不写文件。

## 测试与验收

1. SDK 使用 Blob、Web ReadableStream、AsyncIterable 打开并保持无修改字节 identity。
2. `writeBlob()` 的内容和格式 MIME type 正确。
3. Media 使用 Blob 和 Web stream 添加、去重、删除通过。
4. SmartArt 节点 ID 在无 `node:crypto` import 的实现中生成。
5. Node 文件路径、stream、writeFile 和 CLI 回归通过。
6. browser bundle 构建成功、无静态 `node:` import，并可导入全部根 API/插件命名空间。
7. 全仓 typecheck、测试、Node 聚合包构建通过。
