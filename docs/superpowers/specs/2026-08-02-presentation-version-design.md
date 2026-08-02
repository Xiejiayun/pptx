# Presentation Runtime `version` Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 实例公开的只读 `version: string` 能力，并让 native document 在 source、actual tarball、Node、browser、TypeScript 与 CLI 表面报告自身发布版本：

- `PptxDocument.version` 返回当前 `@jiayunxie/pptx` 发布版本；
- `PPTX_VERSION` 是 root 可导入的同源只读常量；
- create/open/import 后的 document 均返回同一值；
- 该值不进入 OOXML，不影响 package bytes、relationships、mutation journal 或 write output；
- package manifests 与编译期常量由测试锁定，版本升级时必须同步更新。

本小项不复制 PptxGenJS class 上的其他 runtime namespace 属性，不实现 output helpers，不改变任何 presentation create/edit codec，也不声称两个库的版本字符串相同。PptxGenJS 4.0.1 返回 `4.0.1`，本库当前返回 `0.1.0`；对等范围是“document 实例可读取自身库版本”的公开能力。

## 2. 方案比较

### 方案 A：运行时读取 `package.json`

可以避免手写字符串，但 `@pptx/sdk` 独立发布、root tsup bundle、browser conditional export 与 installed CLI 会落在不同目录。JSON module assertion、bundler inline 和文件复制规则会增加不必要的跨运行时分支，也让浏览器 bundle 依赖包布局。

### 方案 B：只导出一个 root 常量

实现最小，但不能对等 PptxGenJS 的 document-instance `version`，调用方仍需知道本库额外的常量命名。

### 方案 C：编译期单一常量 + document getter（采用）

在 SDK 增加 `PPTX_VERSION = '0.1.0' as const`，由 `PptxDocument.version` getter 返回该常量，并通过 SDK/root 导出。测试读取仓库 manifests 锁定同步；发布产物只包含编译期字符串，无运行时文件系统或 bundler JSON 依赖。

## 3. 公共 API

```ts
export const PPTX_VERSION = '0.1.0' as const;
export type PptxVersion = typeof PPTX_VERSION;

export class PptxDocument extends PresentationModel {
  get version(): PptxVersion {
    return PPTX_VERSION;
  }
}
```

典型调用：

```ts
import { PPTX_VERSION, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
console.log(document.version); // '0.1.0'
console.log(document.version === PPTX_VERSION); // true
```

Getter 没有 setter。类型是当前 literal string，调用方可把它当作 string 使用，但不能为 document 赋新版本。`PPTX_VERSION` 使用项目现有常量命名风格，避免与 package-manager metadata 或其他依赖的 generic `version` export 冲突。

## 4. Source of truth 与发布同步

运行时代码不读取文件。仓库测试必须比较：

- root `package.json`；
- `packages/pptx/package.json`；
- `packages/sdk/package.json`；
- `PPTX_VERSION`。

四者必须完全相同并满足稳定 semver `MAJOR.MINOR.PATCH`。这使版本升级成为显式、可 review 的原子变更；任何只修改 manifest 或只修改常量的发布准备都会在测试中失败。

Actual npm tarball smoke 另从 installed root package manifest 读取版本，并要求 installed Node、browser conditional export 与 CLI runtime 报告同一值，防止 source tests 通过但 bundle 内常量陈旧。

## 5. Runtime 与 mutation 语义

- create、open、PptxGenJS import 与六格式 document 都共享同一个 getter；
- getter 是 pure constant read，不解析 package，也不触发 lazy cache；
- 重复读取返回同一 primitive string；
- slide/shape/table/media 编辑、write/reopen 与 transaction rollback 不改变版本；
- version 不写入 core/app properties，也不从文件的 `AppVersion`、producer、extension 或 filename 推断；
- 打开第三方 PPTX 仍报告当前 library runtime version，而不是文件 producer version。

## 6. PptxGenJS 对等边界

PptxGenJS 4.0.1 的 typed public surface 是 `readonly version: string`。Public runtime probe 只验证其实例返回 `4.0.1` 且 create/write 不改变该值。Native 对应验证：

- `document.version === PPTX_VERSION`；
- current native manifest 为 `0.1.0`；
- create/open/write/reopen 前后保持；
- assignment 在 TypeScript 中报错，runtime getter 没有 setter；
- 不要求 native string 等于 PptxGenJS string。

## 7. 验证策略

### Source 与 public API

- 版本常量与三个 manifests 同步；
- `PptxDocument.create()` 与 `open()` 返回 literal version；
- root package 导出常量、类型和 instance getter；
- `@ts-expect-error` 锁定 document assignment；
- getter descriptor 无 setter，读取不产生 package mutation。

### Packed runtime

- actual tarball Node import 比较 installed `package.json` 与 `PPTX_VERSION` / `document.version`；
- browser conditional export 创建 document 并报告 `presentationVersion: true`；
- declaration consumer 接受 `PptxVersion` 与 literal constant，拒绝 assignment；
- installed CLI 的 version/doctor 或 smoke 输出同时带 runtime version evidence。

### Release gates

- focused/full Vitest、performance 与 TypeScript checks；
- Node/browser tsup 与 declaration build；
- actual tarball file inventory、Node/types/browser/CLI 和真实 Chrome；
- `git diff --check`、文档与 compatibility matrix 收尾。

本小项不改变 PPTX bytes，因此不重复生成视觉 gallery 或宣称客户端渲染证据；已有 create/open/write validation 只用于证明 getter 不影响文件路径。

## 8. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. source、SDK、root、actual tarball、Node、browser、types 与 CLI 报告一致 native version；
2. instance getter 与常量公开且只读，manifest drift 有永久测试；
3. create/open/write/reopen 和 package mutation isolation 通过；
4. PptxGenJS public instance behavior 有 conformance evidence，文档明确字符串不相等是正确行为；
5. 全量 release gates 通过，变更 review、commit、push 并确认远端同步；
6. compatibility matrix 将 `version` 从部分支持移入已支持，但继续保留其他 runtime/output helpers 缺口。
