import PptxGenJS from 'pptxgenjs';
import { importPptxGenJS } from '@pptx/pptxgenjs-adapter';

const generated = new PptxGenJS();
generated.addSlide().addText('Hello', { x: 1, y: 1, w: 5, h: 1 });

const document = await importPptxGenJS(generated);
document.slides[0].background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};
await document.writeFile('output.pptx');

