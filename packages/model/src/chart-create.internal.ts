import { relativeRelationshipTarget } from '@pptx/opc';
import type {
  AddChartOptions,
  ChartDefinition,
  ChartGroupInput,
} from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import {
  normalizeAddChartOptions,
  renderChartGraphicFrame,
  renderChartPart,
  type NormalizedAddChartOptions,
} from './chart-render.internal.js';
import {
  buildChartWorkbook,
  planChartWorkbook,
  type ChartWorkbookFormula,
} from './chart-workbook.internal.js';
import { ModelParseError } from './errors.js';
import { resolvePlaceholderOwner } from './placeholder.internal.js';
import { ChartModel } from './shapes.js';
import {
  allocatePresetShapeId,
  directElementChildren,
  requirePresetShapeTree,
  type SlideModel,
} from './slide.js';

const CHART_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const PACKAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export interface PreparedChartCreation {
  readonly definition: Readonly<ChartDefinition>;
  readonly options: Readonly<NormalizedAddChartOptions>;
  readonly workbookBytes: Uint8Array;
  readonly formulas: readonly Readonly<ChartWorkbookFormula>[];
}

export async function prepareChartCreation(
  groups: readonly ChartGroupInput[],
  options?: AddChartOptions,
): Promise<PreparedChartCreation> {
  const definition = normalizeChartDefinition({ groups });
  const normalizedOptions = normalizeAddChartOptions(options);
  const plan = planChartWorkbook(definition);
  const workbookBytes = await buildChartWorkbook(definition);
  return Object.freeze({
    definition,
    options: normalizedOptions,
    workbookBytes,
    formulas: plan.formulas,
  });
}

export function commitPreparedChart(
  slide: SlideModel,
  prepared: PreparedChartCreation,
): ChartModel {
  const pkg = slide.presentation.opcPackage;
  const owner = prepared.options.placeholder === undefined
    ? undefined
    : resolvePlaceholderOwner(
        pkg,
        slide.partUri,
        prepared.options.placeholder,
        'chart',
      );
  const renderedOptions = owner === undefined
    ? prepared.options
    : Object.freeze({
        ...prepared.options,
        name: owner.name,
        ...owner.transform,
      });
  const chartPartUri = pkg.allocatePartUri('/ppt/charts', 'chart', '.xml');
  const workbookPartUri = pkg.allocatePartUri(
    '/ppt/embeddings',
    'Microsoft_Excel_Worksheet',
    '.xlsx',
  );
  pkg.setPart(workbookPartUri, prepared.workbookBytes, WORKBOOK_CONTENT_TYPE);
  const workbookRelationshipId = pkg.allocateRelationshipId(chartPartUri);
  pkg.setPart(
    chartPartUri,
    renderChartPart(prepared.definition, prepared.formulas, workbookRelationshipId),
    CHART_CONTENT_TYPE,
  );
  pkg.addRelationship(chartPartUri, {
    id: workbookRelationshipId,
    type: PACKAGE_RELATIONSHIP_TYPE,
    target: relativeRelationshipTarget(chartPartUri, workbookPartUri),
    targetMode: 'Internal',
  });
  const chartRelationship = pkg.addRelationship(slide.partUri, {
    type: CHART_RELATIONSHIP_TYPE,
    target: relativeRelationshipTarget(slide.partUri, chartPartUri),
    targetMode: 'Internal',
  });
  const { xml } = slide.parse();
  const shapeTree = requirePresetShapeTree(xml, slide.partUri);
  const nextId = owner?.shapeId ?? allocatePresetShapeId(xml, shapeTree, slide.partUri);
  const frame = renderChartGraphicFrame(
    nextId,
    chartRelationship.id,
    renderedOptions,
    owner?.identity,
  );
  if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, frame);
  else {
    const extensionList = directElementChildren(shapeTree, 'extLst')[0];
    if (extensionList) xml.replace(extensionList.start, extensionList.start, frame);
    else xml.appendChildXml(shapeTree, frame);
  }
  slide.setXml(xml.serialize());
  if (owner) slide.invalidateShapeModel(nextId);
  const chart = slide.shapes.find((candidate) => candidate.id === nextId);
  if (!(chart instanceof ChartModel) || chart.kind !== 'chart') {
    throw new ModelParseError(`Created chart ${nextId} could not be resolved`, slide.partUri);
  }
  return chart;
}
