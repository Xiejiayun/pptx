import {
  EMU_PER_POINT,
  type Emu,
  type OoxmlAngle,
  type ShapeLine,
  type SlideSize,
  type Transform,
} from '@pptx/model';

export interface DeckPoint {
  readonly x: Emu;
  readonly y: Emu;
}

export interface DeckRect extends DeckPoint {
  readonly width: Emu;
  readonly height: Emu;
}

export interface DeckInsets {
  readonly top: Emu;
  readonly right: Emu;
  readonly bottom: Emu;
  readonly left: Emu;
}

export type LayoutFamily =
  | 'cinematic-cover'
  | 'section-divider'
  | 'statement-plus-hero'
  | 'asymmetric-two-column'
  | 'large-statistic'
  | 'comparison'
  | 'process-or-timeline'
  | 'chart-or-table'
  | (string & {});

export interface LayoutRegionSpec {
  readonly id: string;
  readonly bounds: DeckRect;
  readonly collision: 'exclusive' | 'overlay' | 'ignore';
  readonly allowBleed?: boolean;
}

interface LayoutElementBase {
  readonly id: string;
  readonly regionId?: string;
  readonly bounds: DeckRect;
  readonly parentId?: string;
  readonly allowBleed?: boolean;
  readonly collision?: 'exclusive' | 'overlay' | 'ignore';
}

export interface TextLayoutSpec extends LayoutElementBase {
  readonly kind: 'text';
  readonly role: 'title' | 'subtitle' | 'body' | 'caption' | 'label';
  readonly text: string;
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly bold?: boolean;
  readonly wrap: boolean;
  readonly maxLines: number;
  readonly fit: 'error' | 'shrink';
  readonly minFontSize: number;
  readonly margin?: Partial<DeckInsets> | Emu;
}

export interface BoxLayoutSpec extends LayoutElementBase {
  readonly kind: 'box' | 'image' | 'chart' | 'table' | 'decoration';
}

export interface ConnectorLayoutSpec {
  readonly kind: 'connector';
  readonly id: string;
  readonly from: DeckPoint;
  readonly to: DeckPoint;
  readonly line?: ShapeLine;
  readonly collision?: 'ignore';
}

export type LayoutElementSpec = TextLayoutSpec | BoxLayoutSpec | ConnectorLayoutSpec;

export interface DeckSlideSpec {
  readonly id: string;
  readonly family: LayoutFamily;
  readonly regions: readonly LayoutRegionSpec[];
  readonly elements: readonly LayoutElementSpec[];
}

export interface DeckSpec {
  readonly schemaVersion: 1;
  readonly slideSize: SlideSize;
  readonly safeArea: DeckInsets;
  readonly gap: Emu;
  readonly fontSafetyFactor?: number;
  readonly slides: readonly DeckSlideSpec[];
}

export type DeckPreflightDiagnosticCode =
  | 'INVALID_GEOMETRY'
  | 'ELEMENT_OUT_OF_BOUNDS'
  | 'SAFE_AREA_VIOLATION'
  | 'REGION_GAP_VIOLATION'
  | 'ELEMENT_OVERLAP'
  | 'TEXT_HORIZONTAL_OVERFLOW'
  | 'TEXT_VERTICAL_OVERFLOW'
  | 'TEXT_MIN_FONT_VIOLATION'
  | 'ZERO_LENGTH_CONNECTOR'
  | 'DUPLICATE_ID';

export interface DeckPreflightDiagnostic {
  readonly severity: 'error' | 'warning';
  readonly code: DeckPreflightDiagnosticCode;
  readonly slideId: string;
  readonly elementId?: string;
  readonly relatedElementId?: string;
  readonly actual?: unknown;
  readonly limit?: unknown;
  readonly suggestion: string;
}

export interface DeckPreflightReport {
  readonly ok: boolean;
  readonly diagnostics: readonly DeckPreflightDiagnostic[];
}

export interface PreflightOptions {
  readonly collisionTolerance?: Emu;
}

const DEFAULT_FONT_SAFETY_FACTOR = 1.1;
const DEFAULT_COLLISION_TOLERANCE = 18_288 as Emu; // 0.02 inch

function finiteRect(rect: DeckRect): boolean {
  return [rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)
    && rect.width > 0
    && rect.height > 0;
}

function contains(outer: DeckRect, inner: DeckRect, tolerance = 0): boolean {
  return inner.x >= outer.x - tolerance
    && inner.y >= outer.y - tolerance
    && inner.x + inner.width <= outer.x + outer.width + tolerance
    && inner.y + inner.height <= outer.y + outer.height + tolerance;
}

function overlap(a: DeckRect, b: DeckRect, tolerance: number): boolean {
  return Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > tolerance
    && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > tolerance;
}

function gapViolation(a: DeckRect, b: DeckRect, gap: number, tolerance: number): boolean {
  if (overlap(a, b, tolerance)) return true;
  const horizontalProjection = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > tolerance;
  const verticalProjection = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > tolerance;
  const horizontalGap = Math.max(b.x - (a.x + a.width), a.x - (b.x + b.width));
  const verticalGap = Math.max(b.y - (a.y + a.height), a.y - (b.y + b.height));
  return (horizontalProjection && horizontalGap < gap - tolerance)
    || (verticalProjection && verticalGap < gap - tolerance);
}

function graphemeWeight(grapheme: string): number {
  if (/^\s$/u.test(grapheme)) return 0.28;
  if (/^[ilI1|'`]$/u.test(grapheme)) return 0.28;
  if (/^[mwMW@%&]$/u.test(grapheme)) return 0.86;
  if (/^[A-Z]$/u.test(grapheme)) return 0.64;
  if (/^[a-z0-9]$/u.test(grapheme)) return 0.53;
  if (/^[.,:;!\-–—()\[\]{}]$/u.test(grapheme)) return 0.36;
  return 1;
}

function textWidthEmu(text: string, fontSize: number, safetyFactor: number): number {
  const graphemes = typeof Intl.Segmenter === 'function'
    ? [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map((part) => part.segment)
    : [...text];
  return graphemes.reduce((sum, grapheme) => sum + graphemeWeight(grapheme), 0)
    * fontSize
    * EMU_PER_POINT
    * safetyFactor;
}

function textMargins(element: TextLayoutSpec): DeckInsets {
  const margin = element.margin;
  if (typeof margin === 'number') {
    return { top: margin, right: margin, bottom: margin, left: margin };
  }
  return {
    top: margin?.top ?? 0 as Emu,
    right: margin?.right ?? 0 as Emu,
    bottom: margin?.bottom ?? 0 as Emu,
    left: margin?.left ?? 0 as Emu,
  };
}

function pushDiagnostic(
  diagnostics: DeckPreflightDiagnostic[],
  diagnostic: Omit<DeckPreflightDiagnostic, 'severity'>,
): void {
  diagnostics.push({ severity: 'error', ...diagnostic });
}

export function connectorTransform(from: DeckPoint, to: DeckPoint): Transform {
  if (![from.x, from.y, to.x, to.y].every(Number.isFinite)) {
    throw new TypeError('Connector endpoints must be finite');
  }
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  if (dx === 0 && dy === 0) throw new RangeError('Connector endpoints must be different');
  return {
    x: Math.min(from.x, to.x) as Emu,
    y: Math.min(from.y, to.y) as Emu,
    width: Math.max(Math.abs(dx), 1) as Emu,
    height: Math.max(Math.abs(dy), 1) as Emu,
    rotation: 0 as OoxmlAngle,
    flipHorizontal: dx < 0,
    flipVertical: dy < 0,
  };
}

export function preflightDeckSpec(
  spec: Readonly<DeckSpec>,
  options: Readonly<PreflightOptions> = {},
): DeckPreflightReport {
  const diagnostics: DeckPreflightDiagnostic[] = [];
  const tolerance = options.collisionTolerance ?? DEFAULT_COLLISION_TOLERANCE;
  const canvas: DeckRect = { x: 0 as Emu, y: 0 as Emu, ...spec.slideSize };
  const safe: DeckRect = {
    x: spec.safeArea.left,
    y: spec.safeArea.top,
    width: (spec.slideSize.width - spec.safeArea.left - spec.safeArea.right) as Emu,
    height: (spec.slideSize.height - spec.safeArea.top - spec.safeArea.bottom) as Emu,
  };

  for (const slide of spec.slides) {
    const ids = new Set<string>();
    const regions = new Map(slide.regions.map((region) => [region.id, region]));
    const boxes = new Map<string, DeckRect>();

    for (const item of [...slide.regions, ...slide.elements]) {
      if (ids.has(item.id)) {
        pushDiagnostic(diagnostics, {
          code: 'DUPLICATE_ID', slideId: slide.id, elementId: item.id,
          suggestion: 'Use a unique id for every region and element on the slide.',
        });
      }
      ids.add(item.id);
    }

    for (const region of slide.regions) {
      if (!finiteRect(region.bounds)) {
        pushDiagnostic(diagnostics, {
          code: 'INVALID_GEOMETRY', slideId: slide.id, elementId: region.id, actual: region.bounds,
          suggestion: 'Use finite coordinates and positive width and height.',
        });
      } else if (!region.allowBleed && !contains(canvas, region.bounds)) {
        pushDiagnostic(diagnostics, {
          code: 'ELEMENT_OUT_OF_BOUNDS', slideId: slide.id, elementId: region.id,
          actual: region.bounds, limit: canvas,
          suggestion: 'Move the region inside the slide or explicitly allow bleed.',
        });
      }
      boxes.set(region.id, region.bounds);
    }
    for (const element of slide.elements) {
      if (element.kind !== 'connector') boxes.set(element.id, element.bounds);
    }

    for (let i = 0; i < slide.regions.length; i += 1) {
      const a = slide.regions[i]!;
      if (a.collision !== 'exclusive' || !finiteRect(a.bounds)) continue;
      for (let j = i + 1; j < slide.regions.length; j += 1) {
        const b = slide.regions[j]!;
        if (b.collision !== 'exclusive' || !finiteRect(b.bounds)) continue;
        if (gapViolation(a.bounds, b.bounds, spec.gap, tolerance)) {
          pushDiagnostic(diagnostics, {
            code: 'REGION_GAP_VIOLATION', slideId: slide.id, elementId: a.id,
            relatedElementId: b.id, actual: 0, limit: spec.gap,
            suggestion: 'Separate exclusive regions by the deck gap.',
          });
        }
      }
    }

    for (const element of slide.elements) {
      if (element.kind === 'connector') {
        try {
          connectorTransform(element.from, element.to);
        } catch {
          pushDiagnostic(diagnostics, {
            code: element.from.x === element.to.x && element.from.y === element.to.y
              ? 'ZERO_LENGTH_CONNECTOR' : 'INVALID_GEOMETRY',
            slideId: slide.id, elementId: element.id,
            suggestion: 'Use two different finite connector endpoints.',
          });
        }
        continue;
      }

      if (!finiteRect(element.bounds)) {
        pushDiagnostic(diagnostics, {
          code: 'INVALID_GEOMETRY', slideId: slide.id, elementId: element.id, actual: element.bounds,
          suggestion: 'Use finite coordinates and positive width and height.',
        });
        continue;
      }
      if (!element.allowBleed && !contains(canvas, element.bounds)) {
        pushDiagnostic(diagnostics, {
          code: 'ELEMENT_OUT_OF_BOUNDS', slideId: slide.id, elementId: element.id,
          actual: element.bounds, limit: canvas,
          suggestion: 'Move the element inside the slide or explicitly allow bleed.',
        });
      } else if (!element.allowBleed && element.kind !== 'decoration' && !contains(safe, element.bounds)) {
        pushDiagnostic(diagnostics, {
          code: 'SAFE_AREA_VIOLATION', slideId: slide.id, elementId: element.id,
          actual: element.bounds, limit: safe,
          suggestion: 'Move audience-facing content inside the safe area.',
        });
      }

      const region = element.regionId ? regions.get(element.regionId) : undefined;
      if (element.regionId && (!region || !contains(region.bounds, element.bounds, tolerance))) {
        pushDiagnostic(diagnostics, {
          code: 'ELEMENT_OVERLAP', slideId: slide.id, elementId: element.id,
          relatedElementId: element.regionId,
          suggestion: 'Keep the element fully inside its declared region.',
        });
      }
      const parentBounds = element.parentId ? boxes.get(element.parentId) : undefined;
      if (element.parentId && (!parentBounds || !contains(parentBounds, element.bounds, tolerance))) {
        pushDiagnostic(diagnostics, {
          code: 'ELEMENT_OVERLAP', slideId: slide.id, elementId: element.id,
          relatedElementId: element.parentId,
          suggestion: 'Keep child content fully inside its parent.',
        });
      }

      if (element.kind !== 'text') continue;
      if (![element.fontSize, element.minFontSize, element.maxLines].every(Number.isFinite)
        || element.fontSize <= 0
        || element.minFontSize <= 0
        || element.maxLines < 1) {
        pushDiagnostic(diagnostics, {
          code: 'INVALID_GEOMETRY', slideId: slide.id, elementId: element.id,
          suggestion: 'Use positive finite font sizes and at least one allowed line.',
        });
        continue;
      }
      const margins = textMargins(element);
      const availableWidth = element.bounds.width - margins.left - margins.right;
      const availableHeight = element.bounds.height - margins.top - margins.bottom;
      if (availableWidth <= 0 || availableHeight <= 0) {
        pushDiagnostic(diagnostics, {
          code: 'INVALID_GEOMETRY', slideId: slide.id, elementId: element.id,
          suggestion: 'Keep text margins smaller than the text box dimensions.',
        });
        continue;
      }
      const safety = spec.fontSafetyFactor ?? DEFAULT_FONT_SAFETY_FACTOR;
      const requiredWidth = textWidthEmu(element.text, element.fontSize, safety);
      const lineCount = element.wrap ? Math.max(1, Math.ceil(requiredWidth / availableWidth)) : 1;
      const requiredHeight = lineCount * element.fontSize * EMU_PER_POINT * 1.14;
      const requiredScale = Math.min(1, availableWidth / requiredWidth, availableHeight / requiredHeight);

      if (lineCount > element.maxLines || (!element.wrap && requiredWidth > availableWidth)) {
        pushDiagnostic(diagnostics, {
          code: 'TEXT_HORIZONTAL_OVERFLOW', slideId: slide.id, elementId: element.id,
          actual: Math.round(requiredWidth), limit: availableWidth,
          suggestion: 'Shorten the copy, widen the box, or explicitly allow more lines.',
        });
      }
      if (requiredHeight > availableHeight || lineCount > element.maxLines) {
        pushDiagnostic(diagnostics, {
          code: 'TEXT_VERTICAL_OVERFLOW', slideId: slide.id, elementId: element.id,
          actual: Math.round(requiredHeight), limit: availableHeight,
          suggestion: 'Shorten the copy or increase the text box height.',
        });
      }
      if (element.fit === 'shrink' && element.fontSize * requiredScale < element.minFontSize) {
        pushDiagnostic(diagnostics, {
          code: 'TEXT_MIN_FONT_VIOLATION', slideId: slide.id, elementId: element.id,
          actual: element.fontSize * requiredScale, limit: element.minFontSize,
          suggestion: 'Shorten the copy or select a roomier layout instead of shrinking further.',
        });
      }
    }
  }

  return { ok: diagnostics.length === 0, diagnostics };
}

export function assertDeckSpec(
  spec: Readonly<DeckSpec>,
  options: Readonly<PreflightOptions> = {},
): void {
  const report = preflightDeckSpec(spec, options);
  if (!report.ok) throw new Error(`DeckSpec preflight failed: ${JSON.stringify(report.diagnostics)}`);
}
