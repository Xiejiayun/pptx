# `@jiayunxie/pptx` 单包发布设计

日期：2026-07-25
状态：已确认

## 目标

将当前 PPTX OOXML monorepo 作为一个可独立安装的公开 npm 包发布：

- 包名：`@jiayunxie/pptx`
- 版本：`0.1.0`
- dist-tag：`next`
- Node.js：20+
- 模块格式：ESM
- 发布渠道：GitHub Actions + npm Trusted Publishing（OIDC）
- CLI：安装后提供 `pptx-inspect`

发布必须不依赖用户无法安装的内部 `@pptx/*` workspace 包，也不使用长期 npm access token。

## 方案选择

采用构建时单包捆绑。使用 `tsup`（以 esbuild 为构建引擎）生成单一 ESM API 入口、CLI 启动文件和类型声明，将 workspace 实现以及 JSZip、Commander 等运行时依赖打入同一个 tarball。

不采用以下方案：

- 重写编译产物 import：JS 与类型声明容易产生路径遗漏。
- 分别发布并重命名所有 workspace 包：增加发布和版本协调成本，且违背单一安装包目标。
- npm `bundledDependencies`：制品中嵌套 `node_modules`，审计和可重复性较差。

## 包结构与公开入口

聚合包位于 `packages/pptx`，`package.json` 名称为 `@jiayunxie/pptx`。

只公开一个 JavaScript/TypeScript 入口：`@jiayunxie/pptx`。

- SDK、model 与内建 codecs 作为根入口的直接命名导出。
- `importPptxGenJS` 作为根入口的直接命名导出。
- 可选能力通过 `transitions`、`animations`、`advancedCharts`、`smartArt` 四个命名空间导出，避免插件之间的符号冲突。
- OPC、lossless XML、validator、testkit 与 CLI 程序化 API 保持包内实现，不承诺独立公开入口。
- `bin.pptx-inspect` 指向捆绑后的 CLI 启动文件，不额外提供 `/cli` 子路径。

package 根目录包含单一根入口对应的 `.d.ts`、README、LICENSE、repository、homepage、bugs、keywords 与 engines 元数据。

## 构建与数据流

1. TypeScript/Vitest 对原 workspace 源码执行现有严格检查。
2. 聚合包以根 API 和 CLI 为两个构建入口，由 esbuild 解析并捆绑所有 workspace import。
3. 类型构建只生成根 API 的公开声明文件，CLI 不暴露程序化类型入口。
4. `pnpm pack` 生成单个 `jiayunxie-pptx-0.1.0.tgz`。
5. 在全新临时项目安装 tarball，验证根 API、四个插件命名空间和 CLI。
6. 只有以上检查全部通过后，GitHub Actions 执行 `npm publish --access public --tag next --provenance`。

## GitHub Actions 与 npm 信任关系

新增 `.github/workflows/publish-npm.yml`，只允许手动触发，并配置：

- `contents: read`
- `id-token: write`
- Node.js 24 与支持 trusted publishing 的 npm CLI
- 安装冻结 lockfile
- 完整检查、构建、pack 和安装 smoke test
- 固定发布 `packages/pptx`，public access，`next` tag，provenance 开启
- GitHub environment：`npm`

npm 侧先创建 staged package `@jiayunxie/pptx`，trusted publisher 精确绑定：

- Owner：`Xiejiayun`
- Repository：`pptx`
- Workflow：`publish-npm.yml`
- Environment：`npm`

不创建或传输 `NPM_TOKEN`。

## 错误处理与安全

- build、test、pack、安装 smoke test 任一失败，publish step 不运行。
- workflow 不接受可改变包名或版本的输入，避免发布任意目标。
- 版本 `0.1.0` 已存在时停止，不自动递增或覆盖。
- trusted publisher、权限或 provenance 校验失败时保留失败日志，不改用长期 token 绕过。
- 不发布 source map 中的本地绝对路径、测试构建、`node_modules`、workspace 协议或私有文件。
- npm/GitHub UI 出现 CAPTCHA、登录验证或需要额外权限时交由用户完成，不绕过验证。

## 测试与验收

发布前必须满足：

- `pnpm check` 全部通过。
- 聚合包构建与 `.d.ts` 生成通过。
- tarball 中无 `workspace:`、`@pptx/*` 运行时依赖、测试文件或 `node_modules`。
- 临时安装后可从根入口 import SDK、adapter 和四个插件命名空间。
- `pptx-inspect --json doctor` 返回成功 JSON，版本为 `0.1.0`。
- GitHub Actions publish job 成功。
- npm 页面显示公开包 `@jiayunxie/pptx@0.1.0`。
- `next` 指向 `0.1.0`，`latest` 不被修改。
- npm 页面显示来自 GitHub Actions 的 provenance。

## 完成定义

当 GitHub Actions 成功发布、npm 包页可访问、安装命令 `npm install @jiayunxie/pptx@next` 可解析，并且 provenance 与 dist-tag 均核对无误时，本任务完成。
