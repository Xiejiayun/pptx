import { PptxDocument, inches } from '@pptx/sdk';

const document = await PptxDocument.open('input.pptx');
document.slides[0].title.text = 'Updated safely';
document.slides[0].shapes[0]?.setTransform({ x: inches(1.25) });
await document.writeFile('output.pptx', {
  compatibility: 'powerpoint-current',
  mode: 'strict',
});

