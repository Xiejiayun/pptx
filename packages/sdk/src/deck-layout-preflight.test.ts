import { describe, expect, it } from 'vitest';
import { inches } from '@pptx/model';
import {
  connectorTransform,
  preflightDeckSpec,
  type DeckSpec,
} from './index.js';

function textDeck(text: string, overrides: Record<string, unknown> = {}): DeckSpec {
  return {
    schemaVersion: 1,
    slideSize: { width: inches(13.333), height: inches(7.5) },
    safeArea: { top: inches(0.4), right: inches(0.4), bottom: inches(0.4), left: inches(0.4) },
    gap: inches(0.3),
    slides: [{
      id: 'slide-7',
      family: 'statement-plus-hero',
      regions: [{
        id: 'title-region',
        bounds: { x: inches(0.62), y: inches(0.42), width: inches(11.9), height: inches(0.64) },
        collision: 'exclusive',
      }],
      elements: [{
        kind: 'text',
        id: 'title',
        regionId: 'title-region',
        bounds: { x: inches(0.62), y: inches(0.42), width: inches(11.9), height: inches(0.64) },
        role: 'title',
        text,
        fontFamily: 'Arial',
        fontSize: 38,
        wrap: false,
        maxLines: 1,
        fit: 'error',
        minFontSize: 36,
        ...overrides,
      }],
    }],
  };
}

describe('deck layout preflight', () => {
  it('catches the clipped Amazon slide 7 title before serialization', () => {
    const report = preflightDeckSpec(textDeck('Deforestation is falling — annual loss remains vast'));
    expect(report.ok).toBe(false);
    expect(report.diagnostics).toContainEqual(expect.objectContaining({
      code: 'TEXT_HORIZONTAL_OVERFLOW',
      slideId: 'slide-7',
      elementId: 'title',
    }));
  });

  it('accepts a shorter one-line title and rejects a box outside the safe area', () => {
    expect(preflightDeckSpec(textDeck('Deforestation leaves lasting scars')).ok).toBe(true);

    const report = preflightDeckSpec(textDeck('Safe title', {
      bounds: { x: inches(0.2), y: inches(0.42), width: inches(3), height: inches(0.64) },
    }));
    expect(report.diagnostics.some((diagnostic) => diagnostic.code === 'SAFE_AREA_VIOLATION')).toBe(true);
  });

  it('normalizes horizontal, vertical, and reverse connectors to positive extents', () => {
    expect(connectorTransform({ x: inches(1), y: inches(2) }, { x: inches(5), y: inches(2) }))
      .toMatchObject({ width: inches(4), height: 1, flipHorizontal: false, flipVertical: false });
    expect(connectorTransform({ x: inches(2), y: inches(1) }, { x: inches(2), y: inches(5) }))
      .toMatchObject({ width: 1, height: inches(4), flipHorizontal: false, flipVertical: false });
    expect(connectorTransform({ x: inches(5), y: inches(5) }, { x: inches(1), y: inches(2) }))
      .toMatchObject({ width: inches(4), height: inches(3), flipHorizontal: true, flipVertical: true });
    expect(() => connectorTransform({ x: inches(1), y: inches(1) }, { x: inches(1), y: inches(1) }))
      .toThrow(/different/);
  });

  it('reports invalid geometry, region collisions, and parent overflow with stable codes', () => {
    const deck = textDeck('Safe title');
    const slide = deck.slides[0]!;
    const report = preflightDeckSpec({
      ...deck,
      slides: [{
        ...slide,
        regions: [
          ...slide.regions,
          {
            id: 'overlap',
            bounds: { x: inches(1), y: inches(0.5), width: inches(2), height: inches(1) },
            collision: 'exclusive',
          },
        ],
        elements: [
          ...slide.elements,
          {
            kind: 'box',
            id: 'bad-box',
            regionId: 'title-region',
            parentId: 'title',
            bounds: { x: Number.NaN as never, y: inches(1), width: inches(1), height: inches(1) },
          },
        ],
      }],
    });
    expect(report.diagnostics.some((diagnostic) => diagnostic.code === 'INVALID_GEOMETRY')).toBe(true);
    expect(report.diagnostics.some((diagnostic) => diagnostic.code === 'REGION_GAP_VIOLATION')).toBe(true);
  });
});
