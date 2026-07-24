# 从 PptxGenJS 接入双向内核

现有 PptxGenJS 创建代码无需迁移到新的对象模型。继续用公开 API 生成演示文稿，再把实例交给 adapter：

```ts
import PptxGenJS from 'pptxgenjs';
import { importPptxGenJS } from '@pptx/pptxgenjs-adapter';

const generated = new PptxGenJS();
generated.addSlide().addText('Hello', { x: 1, y: 1, w: 4, h: 1 });

const document = await importPptxGenJS(generated);
document.slides[0].title.text = 'Edited after generation';
document.duplicateSlide(0);
await document.writeFile('output.pptx');
```

## 依赖边界

`pptxgenjs:^4.0.1` 只由 `@pptx/pptxgenjs-adapter` 直接依赖。adapter 调用 `write({ outputType: 'uint8array' })`；不访问 `_slides`、内部 relationship 计数器或其他私有字段。

## 选择入口

- 从代码新建：PptxGenJS → `importPptxGenJS()` → 双向内核。
- 修改已有文件：直接 `PptxDocument.open()`，不要先导入 PptxGenJS。
- 只需新建且不需要后处理：可继续单独使用 PptxGenJS。

