import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { inches, PptxDocument } from '../../packages/pptx/dist/index.js';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const ASSETS = path.join(DIR, 'assets');
const OUTPUT = path.join(DIR, 'amazon-biodiversity.pptx');

const W = 13.333;
const H = 7.5;
const C = {
  ink: '0A251B',
  forest: '123C2B',
  leaf: '2F6B45',
  fern: '75A66C',
  lime: 'C9E27D',
  mist: 'EEF4E8',
  cream: 'F6F0E2',
  amber: 'E7A93D',
  earth: '8C4D32',
  river: '55B7C4',
  pink: 'E78AA8',
  white: 'FFFFFF',
};

const I = inches;
const srgb = (value) => ({ kind: 'srgb', value });
const fill = (value, transparency) => ({
  kind: 'solid', color: srgb(value),
  ...(transparency === undefined ? {} : { transparency }),
});
const noLine = { kind: 'none' };

function rich(slide, runs, options = {}) {
  return slide.addRichText([{ runs: runs.map((run) => ({
    text: run.text,
    style: {
      fontFamily: run.fontFamily ?? 'Aptos',
      fontSize: run.fontSize ?? 18,
      ...(run.bold === undefined ? {} : { bold: run.bold }),
      ...(run.italic === undefined ? {} : { italic: run.italic }),
      color: srgb(run.color ?? C.ink),
      ...(run.transparency === undefined ? {} : { transparency: run.transparency }),
    },
  })) }], {
    x: I(options.x ?? 0), y: I(options.y ?? 0),
    width: I(options.w ?? 1), height: I(options.h ?? 0.5),
    align: options.align ?? 'left',
    valign: options.valign ?? 'middle',
    margin: options.margin ?? 0,
    fit: options.fit ?? 'shrink',
    wrap: true,
    ...(options.rotation === undefined ? {} : { rotation: options.rotation }),
  });
}

function label(slide, text, x, y, w, color = C.lime) {
  return rich(slide, [{ text: text.toUpperCase(), fontSize: 10, bold: true, color }], {
    x, y, w, h: 0.25,
  });
}

function title(slide, text, x = 0.72, y = 0.55, w = 11.8, color = C.ink) {
  return rich(slide, [{ text, fontFamily: 'Aptos Display', fontSize: 35, bold: true, color }], {
    x, y, w, h: 0.82,
  });
}

function box(slide, x, y, w, h, color, transparency = 0, radius = true) {
  return slide.addShape(radius ? 'roundRect' : 'rect', {
    x: I(x), y: I(y), width: I(w), height: I(h),
    fill: fill(color, transparency), line: noLine,
  });
}

function line(slide, x, y, w, h, color = C.fern, width = 1.5) {
  const normalizedX = w < 0 ? x + w : x;
  const normalizedY = h < 0 ? y + h : y;
  return slide.addShape('line', {
    x: I(normalizedX), y: I(normalizedY),
    width: I(Math.max(Math.abs(w), 0.001)),
    height: I(Math.max(Math.abs(h), 0.001)),
    ...(w < 0 ? { flipHorizontal: true } : {}),
    ...(h < 0 ? { flipVertical: true } : {}),
    line: { kind: 'line', color: srgb(color), width },
  });
}

function numberSlide(slide, n, dark = false) {
  rich(slide, [{ text: String(n).padStart(2, '0'), fontSize: 9, bold: true, color: dark ? C.mist : C.leaf }], {
    x: 12.35, y: 7.02, w: 0.38, h: 0.2, align: 'right',
  });
}

function notes(slide, sources) {
  slide.addNotes(['[Sources]', ...sources].join('\n'));
}

async function image(document, slideIndex, file, x, y, w, h, altText, rounding = false) {
  return document.addImage(slideIndex, path.join(ASSETS, file), {
    x: I(x), y: I(y),
    sizing: { type: 'cover', width: I(w), height: I(h) },
    altText,
    rounding,
  });
}

const document = PptxDocument.create({
  author: 'Codex with @jiayunxie/pptx',
  company: 'OpenAI',
  subject: 'Amazon rainforest biodiversity, ecosystem role, and deforestation threats',
  title: 'Amazon: The Living Climate Engine',
  slideSize: 'wide',
});

// 1 — Cover
{
  const slide = document.addSlide();
  await image(document, 0, 'canopy.jpg', 0, 0, W, H, 'Aerial view of Amazon rainforest crossed by winding brown rivers');
  box(slide, 0, 0, W, H, C.ink, 42, false);
  box(slide, 0, 0, W, 0.16, C.ink, 0, false);
  box(slide, 0, 7.3, W, 0.2, C.ink, 0, false);
  box(slide, 0, 0, 0.18, H, C.lime, 0, false);
  label(slide, 'Biodiversity / climate / continuity', 0.82, 0.75, 4.5, C.lime);
  rich(slide, [
    { text: 'AMAZON', fontFamily: 'Aptos Display', fontSize: 50, bold: true, color: C.white },
  ], { x: 0.78, y: 1.15, w: 7.6, h: 0.95 });
  rich(slide, [
    { text: 'The living climate engine', fontFamily: 'Aptos Display', fontSize: 28, bold: true, color: C.white },
  ], { x: 0.82, y: 2.06, w: 7.8, h: 0.65 });
  rich(slide, [
    { text: 'A forest where species, water and atmosphere are one connected system.', fontSize: 18, color: C.mist },
  ], { x: 0.82, y: 2.92, w: 5.6, h: 0.95, valign: 'top' });
  line(slide, 0.82, 6.47, 1.1, 0, C.lime, 3);
  rich(slide, [{ text: 'A visual field guide', fontSize: 11, bold: true, color: C.mist }], {
    x: 2.08, y: 6.32, w: 2.2, h: 0.3,
  });
  notes(slide, [
    'WWF — About the Amazon: https://wwf.panda.org/discover/knowledge_hub/where_we_work/amazon/about_the_amazon/',
    'Image: canopy.jpg — lubasi, CC BY-SA 2.0; see sources.txt.',
  ]);
}

// 2 — Biodiversity at a glance
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.cream) };
  label(slide, '01 / Scale', 0.72, 0.42, 2.2, C.leaf);
  title(slide, 'Biodiversity is the Amazon’s operating system', 0.72, 0.76, 11.5, C.ink);

  rich(slide, [{ text: '~10%', fontFamily: 'Aptos Display', fontSize: 52, bold: true, color: C.leaf }], {
    x: 0.72, y: 1.62, w: 3.3, h: 1.05,
  });
  rich(slide, [{ text: 'of the world’s known species are associated with the Amazon.', fontSize: 17, color: C.ink }], {
    x: 0.78, y: 2.58, w: 3.55, h: 1.05, valign: 'top',
  });
  box(slide, 0.72, 4.18, 3.62, 1.55, C.forest, 0, true);
  rich(slide, [{ text: '6.7 million km²', fontSize: 25, bold: true, color: C.white }], {
    x: 1.0, y: 4.45, w: 3.05, h: 0.45,
  });
  rich(slide, [{ text: 'A biome roughly twice the size of India.', fontSize: 13, color: C.mist }], {
    x: 1.0, y: 4.95, w: 2.85, h: 0.52, valign: 'top',
  });

  const layers = [
    ['EMERGENT', 'sun, wind, giant crowns', C.lime, 0.9],
    ['CANOPY', 'most leaves, flowers and animal life', C.fern, 1.45],
    ['UNDERSTORY', 'shade-adapted plants and predators', C.leaf, 1.05],
    ['FOREST FLOOR', 'fungi, detritivores and rapid recycling', C.forest, 0.8],
  ];
  let y = 1.68;
  for (const [name, desc, color, h] of layers) {
    box(slide, 5.08, y, 7.45, h, color, 0, false);
    rich(slide, [{ text: name, fontSize: 13, bold: true, color: color === C.lime ? C.ink : C.white }], {
      x: 5.4, y: y + 0.15, w: 1.8, h: 0.3,
    });
    rich(slide, [{ text: desc, fontSize: 13, color: color === C.lime || color === C.fern ? C.ink : C.mist }], {
      x: 7.35, y: y + 0.15, w: 4.7, h: Math.min(h - 0.2, 0.65),
    });
    y += h + 0.07;
  }
  rich(slide, [{ text: 'Four layers. Millions of interactions. One interdependent system.', fontSize: 12, bold: true, color: C.leaf }], {
    x: 5.12, y: 6.58, w: 7.1, h: 0.3, align: 'right',
  });
  numberSlide(slide, 2);
  notes(slide, [
    'WWF — About the Amazon: biome extent and ~10% known-species context: https://wwf.panda.org/discover/knowledge_hub/where_we_work/amazon/about_the_amazon/',
    'Forest-layer descriptions are qualitative synthesis for this presentation.',
  ]);
}

// 3 — Mammals
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.ink) };
  await image(document, 2, 'jaguar.jpg', 0, 0, 7.35, H, 'Jaguar facing forward while walking through forest vegetation');
  box(slide, 0, 0, 7.35, H, C.ink, 32, false);
  await image(document, 2, 'sloth.jpg', 8.55, 0.72, 4.0, 2.35, 'Maned sloth resting in the tree canopy', true);
  box(slide, 6.65, 0, 0.7, H, C.ink, 0, false);
  label(slide, '02 / Mammals', 0.72, 0.48, 2.2, C.lime);
  title(slide, 'Three ways to inhabit a forest', 0.72, 0.82, 6.1, C.white);
  box(slide, 0.72, 4.82, 5.82, 1.58, C.forest, 10, true);
  rich(slide, [{ text: 'JAGUAR', fontSize: 12, bold: true, color: C.amber }], { x: 1.0, y: 5.02, w: 1.5, h: 0.25 });
  rich(slide, [{ text: 'Apex predator', fontSize: 25, bold: true, color: C.white }], { x: 1.0, y: 5.28, w: 2.8, h: 0.42 });
  rich(slide, [{ text: 'Keeps food webs dynamic across vast connected ranges.', fontSize: 13, color: C.mist }], {
    x: 3.55, y: 5.04, w: 2.55, h: 0.82, valign: 'top',
  });

  rich(slide, [{ text: 'SLOTH', fontSize: 11, bold: true, color: C.lime }], { x: 8.55, y: 3.28, w: 1.0, h: 0.25 });
  rich(slide, [{ text: 'A moving microhabitat', fontSize: 20, bold: true, color: C.white }], { x: 8.55, y: 3.55, w: 4.0, h: 0.45 });
  rich(slide, [{ text: 'Slow metabolism, canopy camouflage and a coat shared with algae and invertebrates.', fontSize: 13, color: C.mist }], {
    x: 8.55, y: 4.05, w: 3.9, h: 0.95, valign: 'top',
  });
  slide.addShape('ellipse', { x: I(8.55), y: I(5.38), width: I(0.7), height: I(0.7), fill: fill(C.pink), line: noLine });
  line(slide, 9.32, 5.73, 0.75, 0, C.pink, 2.5);
  rich(slide, [{ text: 'PINK RIVER DOLPHIN', fontSize: 11, bold: true, color: C.pink }], { x: 10.2, y: 5.47, w: 2.2, h: 0.25 });
  rich(slide, [{ text: 'Navigates flooded forest by sound.', fontSize: 14, color: C.white }], { x: 10.2, y: 5.82, w: 2.2, h: 0.6, valign: 'top' });
  numberSlide(slide, 3, true);
  notes(slide, [
    'WWF — Jaguar: https://www.worldwildlife.org/species/jaguar',
    'WWF — Amazon River Dolphin: https://www.worldwildlife.org/species/amazon-river-dolphin',
    'Pauli et al. (2014) — sloth fur micro-ecosystem: https://doi.org/10.1098/rspb.2013.3006',
    'Image: jaguar.jpg — Charles J. Sharp, CC BY-SA 4.0; see sources.txt.',
    'Image: sloth.jpg — deboas, CC BY 4.0; see sources.txt.',
  ]);
}

// 4 — Insects
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.mist) };
  await image(document, 3, 'ants.jpg', 6.15, 0, 7.18, H, 'Leafcutter ant carrying a leaf segment along tree bark');
  box(slide, 5.72, 0, 1.25, H, C.mist, 0, false);
  label(slide, '03 / Invertebrates', 0.72, 0.46, 2.6, C.leaf);
  title(slide, 'The small species do the heavy work', 0.72, 0.82, 5.3, C.ink);
  rich(slide, [{ text: 'Leafcutter ants are not leaf eaters.', fontSize: 19, bold: true, color: C.earth }], {
    x: 0.76, y: 1.7, w: 4.85, h: 0.6,
  });
  rich(slide, [{ text: 'They harvest leaves to cultivate fungus—the colony’s food supply.', fontSize: 15, color: C.ink }], {
    x: 0.76, y: 2.3, w: 4.65, h: 0.8, valign: 'top',
  });
  const functions = [
    ['CUT', 'fresh vegetation', C.amber],
    ['CULTIVATE', 'fungus gardens', C.fern],
    ['HARVEST', 'fungus as colony food', C.leaf],
  ];
  let y = 3.45;
  functions.forEach(([verb, desc, color], i) => {
    slide.addShape('ellipse', { x: I(0.78), y: I(y), width: I(0.62), height: I(0.62), fill: fill(color), line: noLine });
    rich(slide, [{ text: String(i + 1), fontSize: 14, bold: true, color: i === 0 ? C.ink : C.white }], {
      x: 0.78, y, w: 0.62, h: 0.62, align: 'center',
    });
    rich(slide, [{ text: verb, fontSize: 13, bold: true, color: C.ink }], { x: 1.62, y: y - 0.02, w: 1.2, h: 0.25 });
    rich(slide, [{ text: desc, fontSize: 14, color: C.leaf }], { x: 2.82, y: y - 0.02, w: 2.35, h: 0.3 });
    if (i < functions.length - 1) line(slide, 1.09, y + 0.67, 0, 0.42, C.fern, 1.5);
    y += 1.05;
  });
  numberSlide(slide, 4, true);
  notes(slide, [
    'Smithsonian National Zoo — Leafcutter Ant: fungus cultivation and colony ecology: https://nationalzoo.si.edu/animals/leafcutter-ant',
    'Image: ants.jpg — Matheysil, CC BY-SA 4.0; see sources.txt.',
  ]);
}

// 5 — Plants
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.forest) };
  await image(document, 4, 'plants.jpg', 0, 0, 7.35, H, 'Monstera dubia climbing a rainforest tree trunk');
  box(slide, 0, 0, 7.35, H, C.ink, 24, false);
  label(slide, '04 / Plants', 0.72, 0.48, 2.2, C.lime);
  title(slide, 'The forest is a vertical city', 0.72, 0.82, 5.8, C.white);
  rich(slide, [{ text: 'Light is the scarce currency.', fontSize: 18, bold: true, color: C.lime }], { x: 0.78, y: 5.15, w: 3.15, h: 0.4 });
  rich(slide, [{ text: 'Trees race upward; lianas borrow trunks; epiphytes build gardens in the canopy.', fontSize: 14, color: C.mist }], {
    x: 0.78, y: 5.62, w: 4.8, h: 0.85, valign: 'top',
  });

  rich(slide, [{ text: '~390', fontFamily: 'Aptos Display', fontSize: 48, bold: true, color: C.lime }], { x: 8.05, y: 1.12, w: 2.4, h: 0.9 });
  rich(slide, [{ text: 'billion trees', fontSize: 16, bold: true, color: C.white }], { x: 10.45, y: 1.5, w: 2.0, h: 0.35 });
  line(slide, 8.05, 2.24, 4.2, 0, C.fern, 1.5);
  rich(slide, [{ text: '~16,000', fontFamily: 'Aptos Display', fontSize: 39, bold: true, color: C.amber }], { x: 8.05, y: 2.6, w: 2.45, h: 0.75 });
  rich(slide, [{ text: 'tree species', fontSize: 16, bold: true, color: C.white }], { x: 10.48, y: 2.89, w: 1.8, h: 0.35 });
  box(slide, 8.0, 4.15, 4.5, 1.55, C.ink, 18, true);
  rich(slide, [{ text: 'Diversity is uneven', fontSize: 20, bold: true, color: C.white }], { x: 8.35, y: 4.38, w: 3.8, h: 0.4 });
  rich(slide, [{ text: 'A small set of “hyperdominant” trees accounts for a large share of stems—yet rare species hold irreplaceable genetic options.', fontSize: 13, color: C.mist }], {
    x: 8.35, y: 4.88, w: 3.72, h: 0.68, valign: 'top',
  });
  numberSlide(slide, 5, true);
  notes(slide, [
    'ter Steege et al., Science (2013) — ~390 billion trees and ~16,000 tree species: https://doi.org/10.1126/science.1243092',
    'Image: plants.jpg — Dick Culbert, CC BY 2.0; see sources.txt.',
  ]);
}

// 6 — Global role
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.ink) };
  label(slide, '05 / Planetary role', 0.72, 0.44, 2.6, C.lime);
  title(slide, 'The Amazon moves water before it stores carbon', 0.72, 0.78, 11.7, C.white);
  rich(slide, [{ text: 'FOREST', fontSize: 12, bold: true, color: C.ink }], { x: 0.78, y: 2.0, w: 1.2, h: 0.26, align: 'center' });
  slide.addShape('ellipse', { x: I(0.78), y: I(2.36), width: I(2.4), height: I(2.4), fill: fill(C.lime), line: noLine });
  rich(slide, [{ text: 'LEAVES', fontSize: 19, bold: true, color: C.ink }], { x: 1.12, y: 2.95, w: 1.72, h: 0.35, align: 'center' });
  rich(slide, [{ text: 'release moisture', fontSize: 14, color: C.ink }], { x: 1.12, y: 3.38, w: 1.72, h: 0.62, align: 'center' });

  line(slide, 3.33, 3.55, 1.22, 0, C.river, 4);
  slide.addShape('ellipse', { x: I(4.7), y: I(2.36), width: I(2.4), height: I(2.4), fill: fill(C.river), line: noLine });
  rich(slide, [{ text: 'AIR', fontSize: 19, bold: true, color: C.ink }], { x: 5.04, y: 2.95, w: 1.72, h: 0.35, align: 'center' });
  rich(slide, [{ text: 'transports vapor', fontSize: 14, color: C.ink }], { x: 5.04, y: 3.38, w: 1.72, h: 0.62, align: 'center' });

  line(slide, 7.25, 3.55, 1.22, 0, C.river, 4);
  slide.addShape('ellipse', { x: I(8.62), y: I(2.36), width: I(2.4), height: I(2.4), fill: fill(C.mist), line: noLine });
  rich(slide, [{ text: 'RAIN', fontSize: 19, bold: true, color: C.ink }], { x: 8.96, y: 2.95, w: 1.72, h: 0.35, align: 'center' });
  rich(slide, [{ text: 'feeds rivers + farms', fontSize: 14, color: C.ink }], { x: 8.96, y: 3.38, w: 1.72, h: 0.62, align: 'center' });

  line(slide, 10.92, 4.85, -8.65, 0.95, C.fern, 2);
  rich(slide, [{ text: 'RECYCLED MOISTURE RETURNS', fontSize: 12, bold: true, color: C.lime }], { x: 4.05, y: 5.5, w: 4.05, h: 0.3, align: 'center' });
  box(slide, 0.78, 6.15, 11.55, 0.58, C.lime, 0, true);
  rich(slide, [{ text: 'At the same time, living biomass locks up carbon—until clearing and fire reverse the flow.', fontSize: 14, bold: true, color: C.ink }], {
    x: 1.05, y: 6.29, w: 11.0, h: 0.3, align: 'center',
  });
  numberSlide(slide, 6, true);
  notes(slide, [
    'NASA Earth Observatory — The Amazon’s Vicious Cycles: moisture recycling, drought, fire, and land-use feedbacks: https://earthobservatory.nasa.gov/features/AmazonFire',
    'IPCC AR6 WGII Chapter 12: https://www.ipcc.ch/report/ar6/wg2/chapter/chapter-12/',
  ]);
}

// 7 — Threats
{
  const slide = document.addSlide();
  await image(document, 6, 'deforestation.jpg', 0, 0, W, H, 'Satellite view of fragmented Amazon forest with red hotspot markers');
  box(slide, 0, 0, W, H, C.ink, 45, false);
  box(slide, 0, 0, 4.65, H, C.ink, 12, false);
  box(slide, 0, 0, W, 0.16, C.ink, 0, false);
  box(slide, 0, 7.3, W, 0.2, C.ink, 0, false);
  label(slide, '06 / Threat', 0.72, 0.48, 2.2, C.amber);
  title(slide, 'Deforestation becomes a feedback loop', 0.72, 0.84, 4.0, C.white);
  rich(slide, [{ text: 'Clearing removes habitat once. Drying and fire can keep removing it.', fontSize: 17, color: C.mist }], {
    x: 0.78, y: 2.28, w: 3.4, h: 1.15, valign: 'top',
  });
  box(slide, 0.78, 4.72, 3.45, 1.25, C.earth, 0, true);
  rich(slide, [{ text: 'The risk is cumulative', fontSize: 20, bold: true, color: C.white }], { x: 1.05, y: 4.93, w: 2.9, h: 0.38 });
  rich(slide, [{ text: 'roads + extraction + heat + fire', fontSize: 12, color: C.cream }], { x: 1.05, y: 5.42, w: 2.9, h: 0.26 });

  const chain = [
    ['CLEAR', 'habitat removed', C.amber],
    ['FRAGMENT', 'edges heat + dry', 'D38A3B'],
    ['DRY', 'edges lose moisture', C.earth],
    ['BURN', 'fire releases carbon', '6B4438'],
    ['LOSE', 'species + resilience', C.ink],
  ];
  let x = 5.2;
  chain.forEach(([verb, desc, color], i) => {
    box(slide, x, 4.45 - i * 0.5, 1.35, 1.45, color, 0, true);
    rich(slide, [{ text: verb, fontSize: 13, bold: true, color: C.white }], { x: x + 0.12, y: 4.62 - i * 0.5, w: 1.11, h: 0.28, align: 'center' });
    rich(slide, [{ text: desc, fontSize: 10, color: C.mist }], { x: x + 0.12, y: 4.98 - i * 0.5, w: 1.11, h: 0.48, align: 'center', valign: 'top' });
    if (i < chain.length - 1) line(slide, x + 1.4, 5.15 - i * 0.5, 0.35, -0.5, C.amber, 2);
    x += 1.55;
  });
  numberSlide(slide, 7, true);
  notes(slide, [
    'NASA Earth Observatory — The Amazon’s Vicious Cycles: https://earthobservatory.nasa.gov/features/AmazonFire',
    'IPCC AR6 WGII Chapter 12: https://www.ipcc.ch/report/ar6/wg2/chapter/chapter-12/',
    'Global Forest Watch — Brazil forest-change dashboard: https://www.globalforestwatch.org/dashboards/country/BRA/?category=forest-change',
    'Image: deforestation.jpg — NASA LANCE FIRMS / ESDIS, public domain; see sources.txt.',
  ]);
}

// 8 — Action
{
  const slide = document.addSlide();
  slide.background = { kind: 'solid', color: srgb(C.cream) };
  label(slide, '07 / Response', 0.72, 0.45, 2.2, C.leaf);
  title(slide, 'Protect continuity—not isolated fragments', 0.72, 0.8, 11.6, C.ink);
  rich(slide, [{ text: 'The strongest strategy keeps ecological processes connected across landscapes and generations.', fontSize: 17, color: C.leaf }], {
    x: 0.75, y: 1.55, w: 8.6, h: 0.7,
  });

  const actions = [
    ['01', 'RIGHTS', 'Secure Indigenous and local-community stewardship.', C.lime],
    ['02', 'ENFORCEMENT', 'Stop illegal clearing, mining and wildlife trade.', C.fern],
    ['03', 'RESTORATION', 'Reconnect riverbanks, corridors and degraded edges.', C.leaf],
    ['04', 'MARKETS', 'Trace commodities and reward standing forest.', C.forest],
  ];
  let y = 2.55;
  actions.forEach(([num, verb, desc, color]) => {
    box(slide, 0.78, y, 11.78, 0.82, color, 0, true);
    const darkText = color === C.lime || color === C.fern;
    rich(slide, [{ text: num, fontSize: 12, bold: true, color: darkText ? C.ink : C.white }], { x: 1.05, y: y + 0.18, w: 0.45, h: 0.28, align: 'center' });
    rich(slide, [{ text: verb, fontSize: 13, bold: true, color: darkText ? C.ink : C.white }], { x: 1.8, y: y + 0.18, w: 1.85, h: 0.28 });
    rich(slide, [{ text: desc, fontSize: 14, color: darkText ? C.ink : C.mist }], { x: 4.05, y: y + 0.16, w: 7.85, h: 0.35 });
    y += 0.95;
  });

  rich(slide, [{ text: 'A living Amazon is infrastructure for life.', fontFamily: 'Aptos Display', fontSize: 23, bold: true, color: C.earth }], {
    x: 0.78, y: 6.52, w: 8.2, h: 0.45,
  });
  numberSlide(slide, 8);
  notes(slide, [
    'IPCC AR6 WGII Chapter 12 — conservation, adaptation, and land-use context: https://www.ipcc.ch/report/ar6/wg2/chapter/chapter-12/',
    'Global Forest Watch — monitoring context: https://www.globalforestwatch.org/dashboards/country/BRA/?category=forest-change',
    'Action framing is a presentation synthesis, not a ranked policy assessment.',
  ]);
}

await document.writeFile(OUTPUT, { compression: true });
console.log(OUTPUT);
