# ADR 0003：兼容 profile 与受控降级

- 状态：Accepted
- 日期：2026-07-25

输出能力分为 `native`、`preserved`、`degraded`、`unsupported`。默认 preserve-first，不自动栅格化、不删除未知内容、不下载外链资源。

首批 profile 为 `powerpoint-2010`、`powerpoint-current`、`keynote-current`、`libreoffice-current` 和 `google-slides-import`。严格模式遇到 error diagnostic 阻止保存；宽松模式输出文件并返回完整 diagnostic。

