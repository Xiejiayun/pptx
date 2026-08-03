import type { XmlElement } from '@pptx/lossless-xml';

export function readDirectTablePhysicalCells(
  frame: XmlElement,
): readonly XmlElement[] | undefined {
  return readDirectTablePhysicalCellMatrix(frame)?.flat();
}

export function readDirectTablePhysicalCellMatrix(
  frame: XmlElement,
): readonly (readonly XmlElement[])[] | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic ? exactDirectChild(graphic, 'graphicData') : undefined;
  const table = graphicData ? exactDirectChild(graphicData, 'tbl') : undefined;
  if (!table) return undefined;
  const rows = directChildren(table, 'tr');
  if (rows.length === 0) return undefined;
  const matrix = rows.map((row) => directChildren(row, 'tc'));
  return matrix.some((cells) => cells.length === 0) ? undefined : matrix;
}

function exactDirectChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
