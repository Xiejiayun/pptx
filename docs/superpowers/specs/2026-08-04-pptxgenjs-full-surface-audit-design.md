# PptxGenJS 4.0.1 Full-Surface Audit Design

日期：2026-08-04
状态：已批准实施

## 目标

建立一个可重复、可审查、默认失败的 PptxGenJS 4.0.1 全公开表面审计系统。它必须从锁定依赖的真实声明与运行时生成权威清单，把每个公开能力原子映射到 native API、实现代码、定向测试、actual-package/browser 证明和 PPTX 结构证据，并把任何缺失、陈旧或间接证据判为未完成。

本审计取代“根据已有文档未发现缺口”或“某个总 smoke 为 true”这类弱结论。只有完整清单中不存在 `unsupported`、`unverified` 或 `stale` 条目，所有合法公开能力都具有直接证据，且最终客户端 corpus 达到明确通过标准时，才允许声明完整 PptxGenJS parity 已认证。

用户已明确授权实现方自行选择最佳方案、持续推进并且不设置询问点，因此本设计在完成方案比较与自审后直接批准。

## 范围基准

唯一版本基准是 workspace lockfile 中的 `pptxgenjs@4.0.1`：

- `types/index.d.ts` 定义公开声明表面；
- `dist/pptxgen.cjs.js` 与真实实例定义公开运行时表面；
- 由公开 API 生成的真实 PPTX 定义合法最终输出语义；
- PptxGenJS 自身的私有字段、未声明 helper、已确认的无效 OOXML、caller mutation、truthy coercion 和其他 runtime 缺陷不构成 native 兼容目标。

对等不要求现有 PptxGenJS 调用只替换 import 后原样运行。Native 可以使用更严格、可逆、事务安全的高层 API，但必须覆盖相同的合法能力，并由真实输出对照证明最终语义没有减少。

## 为什么现有结论不足

当前 baseline 表格覆盖了许多主能力，但同一文件仍明确列出 percentage coordinates、完整 theme text cascade、custom shows、放映范围、advanced shape/image/media/chart style、hover/action links 等未完成项。最后的“公开能力覆盖 100%”只针对此前人工锁定的能力列表，不能证明整个 `index.d.ts` 已逐属性闭合。

现有 full test、packed smoke 和 Chrome smoke 是重要发布证据，但它们没有自动证明：

- 声明新增或遗漏的 option property 已进入审计；
- 每个公开 union token 都有正向与非法边界覆盖；
- 同名能力在 slide/master/layout/placeholder/table 等不同 owner 上均已覆盖；
- 文档中的“deliberate difference”确实保持合法能力，而不是缩小范围；
- 客户端 corpus 覆盖了所有需要客户端验证的对象类别。

因此需要 declaration → manifest → evidence 的闭环，而不是继续维护一个无法自动防漂移的百分比。

## 方案比较

### 方案一：继续人工维护 Markdown 矩阵

优点是改动少，缺点是无法发现声明漂移、同名字段遗漏、无效证据路径和重复计数。当前矛盾正是这种模式产生的，因此不采用。

### 方案二：只解析 TypeScript 声明并比较名称

它能得到完整符号和属性名称，但名称相同不代表行为对等，也无法区分合法 strict difference、PptxGenJS 缺陷和 native extension。单独使用会产生大量错误完成结论，因此不采用。

### 方案三：声明、运行时、真实 OOXML 三源审计

采用。TypeScript AST 生成稳定公开清单；真实运行时探针锁定 catalogs、实例属性和关键有效调用；人工审查的 manifest 只负责把能力原子映射到 native 设计与直接证据；真实 PptxGenJS/native 输出和 `pptx-inspect` 负责最终语义证明。任何一层缺失都保持未完成。

## 审计原子

审计器从以下 public roots 出发：

- default `PptxGenJS` class 的 public property、getter 和 method；
- namespace `PptxGenJS.Slide` 的 public property 和 method；
- 这些方法直接引用的 option/data interfaces；
- public catalogs 与 string-literal unions，例如 chart、shape、placeholder、layout、alignment、output、scheme color；
- option interface 的每个 own property，包括继承链展开后的来源；
- nested public record 的每个 property，只要它可从上述 root 的合法输入到达。

每个原子使用稳定 ID：

```text
class:PptxGenJS#write
class:PptxGenJS@property:theme
interface:TextPropsOptions@property:margin
interface:ShapeProps@property:shadow
union:OUTPUT_TYPE#nodebuffer
method:Slide#addMedia
```

函数重载共享一个 method ID，并在 declaration metadata 中保存全部 signatures。Interface inheritance 被展开，但保留 `declaredIn`，避免同一继承字段被重复计数。Internal helper type 只有在 public root 可达时才进入清单。

## 状态模型

每个原子必须且只能属于以下状态之一：

- `supported`：合法公开能力有 native 高层 API，且代码、定向测试、consumer/制品和 OOXML 证据完整；
- `deliberate-difference`：native API 形状或严格性不同，但覆盖相同合法能力，真实 PptxGenJS 对照和差异测试完整；
- `deprecated-alias`：canonical PptxGenJS 能力已覆盖，alias 不需要复制，但必须记录 canonical target 和真实 alias 行为；
- `defect-excluded`：只代表 PptxGenJS 无效输出、caller mutation、coercion 或未声明行为，不计入能力闭合；
- `unsupported`：确认存在合法公开能力缺口；
- `unverified`：可能已实现，但证据不足；
- `stale`：manifest 引用的 declaration atom、文件、测试或证据已经不存在或不匹配。

Native extension 单独记录为 `extensions`，永远不能抵消一个 PptxGenJS `unsupported` 原子，也不进入完成率分母。

完整 parity 的 declaration gate 要求 `unsupported = 0`、`unverified = 0`、`stale = 0`。`deliberate-difference`、`deprecated-alias` 和 `defect-excluded` 必须各自具有直接基线测试，不能只靠说明文字关闭。

## 文件与模块边界

### `scripts/pptxgenjs-surface-audit.mjs`

负责定位锁定依赖、使用 workspace TypeScript compiler 解析 `types/index.d.ts`、提取稳定清单、加载 manifest、验证证据引用并输出 JSON/Markdown summary。它不修改源代码，也不调用网络。

依赖定位从 `packages/pptxgenjs-adapter/package.json` 创建 `require` context，再解析 `pptxgenjs/package.json`，不能硬编码 `.pnpm` store 路径。

### `scripts/pptxgenjs-surface-manifest.mjs`

导出只含 plain data 的 manifest。每项包含：

```js
{
  id: 'interface:ImageProps@property:transparency',
  status: 'unsupported',
  native: [],
  evidence: {
    code: [],
    tests: [],
    package: [],
    ooxml: [],
    clients: [],
  },
  note: 'PptxGenJS 4.0.1 legal percentage transparency has no native image API.',
}
```

Manifest 不复制 declaration 名称列表；缺少 manifest entry 的 declaration atom 自动成为 `unverified`。Manifest 中不存在于最新 declaration 的 entry 自动成为 `stale`。

### `scripts/pptxgenjs-runtime-probe.mjs`

只通过 PptxGenJS 公开 API 创建真实实例，稳定输出：own/prototype public keys、runtime catalogs、合法最小 presentation/slide calls 的返回类型，以及需要区分声明与 runtime 的 canonical/deprecated 行为。探针不得读取 `_slides` 等私有字段来证明能力。

### `docs/compatibility/pptxgenjs-surface-audit.json`

由审计器生成的稳定制品，包含 dependency hash、declaration atom 数、各状态计数、未完成 ID、manifest/evidence diagnostics 和 runtime probe hash。排序固定，不包含绝对路径、时间戳或机器相关值。

### `docs/compatibility/pptxgenjs-surface-audit.md`

面向 reviewer 的生成摘要，按 presentation、slide lifecycle、text、shape、image、media、chart、table、master/layout、output/runtime 分组。它只展示结果和证据链接，不成为权威输入。

## 证据规则

`supported` 至少需要：

1. 一个 native public API 或明确的 import/edit path；
2. 一个精确到 test file 和 test title pattern 的定向测试；
3. 一个 actual packed consumer 或声明/浏览器 smoke 证据；
4. 对会序列化到 PPTX 的能力，至少一个 `pptx-inspect` part/relationship/validation assertion；
5. 对视觉或客户端相关能力，至少一个 render/overflow/client corpus assertion。

纯 runtime catalog、metadata 或返回类型不要求 OOXML 证据，但仍要求 declarations、Node/browser consumer 和 mutation isolation。纯 import-preservation 能力必须有 before/after package diff，而不能用创建测试替代。

审计器验证所有 repository-relative evidence path 存在，测试 pattern 能在目标文件中找到，commit ID 能由当前 Git object database 解析。`/tmp` evidence 只可作为补充，不能成为唯一长期证据。

## 真实输出对照

对每个需要语义确认的 gap family，先由锁定 PptxGenJS 4.0.1 生成最小合法 control deck，再由 native API 生成或导入编辑等价 deck。对照重点是 capability-level final state：

- content types、parts 和 relationships；
- object type、geometry、style、text、data 和 metadata；
- duplicate/move/delete/reopen 与 failure isolation；
- PowerPoint 2010 profile diagnostics；
- render/overflow 和可用客户端的保存行为。

Native 可以写出更合法、更明确或更可逆的 OOXML，但必须在 manifest 中记录差异，测试必须证明没有丢失合法能力。

## 实施分解

本设计只覆盖“全量审计基础设施与第一份可信缺口矩阵”。审计完成后，每个独立 gap family 都建立自己的设计、计划、实现和证明 commit，不把多个无关 subsystem 合并成一个大改动。

第一阶段顺序：

1. declaration extractor 与稳定 atom IDs；
2. runtime probe 与 dependency/hash 防漂移；
3. manifest schema、evidence verifier 和默认失败 gate；
4. 初始 manifest 映射与生成 JSON/Markdown；
5. reviewer 复核所有 `unsupported` / `unverified`，形成后续小项队列；
6. 将 audit gate 加入 full/package smoke，但在矩阵完成前明确失败而不是伪造 100%。

随后按依赖和用户价值优先实现：shared position/metadata → shape/text advanced styles → image lifecycle/styles → slide presentation controls → media/chart advanced public options → master/layout cascade → 客户端 corpus。

## 测试与验收

- 锁定 4.0.1 declaration hash 和 package version；版本或 hash 改变时审计失败；
- AST fixture 覆盖 overload、interface inheritance、optional/readonly property、nested records、unions 和 recursive references；
- 生成两次 JSON/Markdown byte-identical；
- 删除一个 manifest entry 时出现 `unverified`；增加不存在的 entry 时出现 `stale`；
- 删除 evidence file、test pattern 或 commit object 时 gate 失败；
- 同一 atom 重复、状态非法、extension 抵消 unsupported 时 gate 失败；
- runtime probe 只使用公开 API，并在 Node 下稳定；
- 初始矩阵必须如实显示现有缺口，不能为了让 gate 通过把未知项批量标为 supported；
- audit 工具本身通过 Node check、focused tests、TypeScript/root build 和实际仓库运行；
- 每个实施小项继续执行 review、commit、push 和远端同步检查。

## 完成标准

本审计基础设施完成时：

- PptxGenJS 4.0.1 public declaration atom 数量和分类可重复生成；
- 每个 atom 都有唯一 manifest 状态或被自动列为 `unverified`；
- 所有 evidence 引用都经过机器验证；
- 生成摘要明确列出真实缺口，不再输出未经证明的 100%；
- 后续实现队列由 `unsupported` / `unverified` 稳定驱动。

总目标完成标准仍更严格：所有合法 public atoms 都被直接证明关闭，所有客户端要求完成，full/package/browser/PPTX/client gates 全部通过，且最终完成性审计没有缺失或间接证据。
