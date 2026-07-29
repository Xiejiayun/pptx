import {
  PRESENTATION_FORMAT_PROFILES,
  type PresentationFormat,
  type SlideSize,
} from '@pptx/model';
import { OpcPackage } from '@pptx/opc';

export type BuiltInSlideSize = '4:3' | '16:9' | '16:10' | 'wide';

export type CustomSlideSize = SlideSize;

export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly title?: string;
}

const SLIDE_SIZES: Readonly<Record<BuiltInSlideSize, { readonly cx: number; readonly cy: number }>> = {
  '4:3': { cx: 9_144_000, cy: 6_858_000 },
  '16:9': { cx: 9_144_000, cy: 5_143_500 },
  '16:10': { cx: 9_144_000, cy: 5_715_000 },
  wide: { cx: 12_192_000, cy: 6_858_000 },
};

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CONTENT = 'application/vnd.openxmlformats-officedocument.presentationml.';
const MIN_SLIDE_SIZE = 914_400;
const MAX_SLIDE_SIZE = 51_206_400;

export function createPresentationPackage(options: CreatePresentationOptions = {}): OpcPackage {
  const format = options.format ?? 'pptx';
  const rtlMode = options.rtlMode;
  const slideSize = options.slideSize === undefined ? '16:9' : options.slideSize;
  if (!Object.hasOwn(PRESENTATION_FORMAT_PROFILES, format)) {
    throw new TypeError(`Unsupported presentation format: ${String(format)}`);
  }
  if (rtlMode !== undefined && typeof rtlMode !== 'boolean') {
    throw new TypeError('Presentation RTL mode must be a boolean');
  }
  const profile = PRESENTATION_FORMAT_PROFILES[format];
  const size = resolveSlideSize(slideSize);
  const pkg = OpcPackage.create();

  return pkg.transaction(() => {
    pkg.setPart('/docProps/core.xml', CORE_PROPERTIES_XML, 'application/vnd.openxmlformats-package.core-properties+xml');
    pkg.setPart('/docProps/app.xml', APP_PROPERTIES_XML, 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
    pkg.setPart(
      '/ppt/presentation.xml',
      presentationXml(size.cx, size.cy, rtlMode),
      profile.presentationContentType,
    );
    pkg.setPart('/ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER_XML, `${CONTENT}slideMaster+xml`);
    pkg.setPart('/ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT_XML, `${CONTENT}slideLayout+xml`);
    pkg.setPart('/ppt/theme/theme1.xml', THEME_XML, 'application/vnd.openxmlformats-officedocument.theme+xml');
    pkg.setPart('/ppt/notesMasters/notesMaster1.xml', NOTES_MASTER_XML, `${CONTENT}notesMaster+xml`);
    pkg.setPart('/ppt/presProps.xml', PRESENTATION_PROPERTIES_XML, `${CONTENT}presProps+xml`);
    pkg.setPart('/ppt/viewProps.xml', VIEW_PROPERTIES_XML, `${CONTENT}viewProps+xml`);
    pkg.setPart('/ppt/tableStyles.xml', TABLE_STYLES_XML, `${CONTENT}tableStyles+xml`);

    pkg.addRelationship('/', {
      id: 'rId1',
      type: `${REL}officeDocument`,
      target: 'ppt/presentation.xml',
    });
    pkg.addRelationship('/', {
      id: 'rId2',
      type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
      target: 'docProps/core.xml',
    });
    pkg.addRelationship('/', {
      id: 'rId3',
      type: `${REL}extended-properties`,
      target: 'docProps/app.xml',
    });

    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId1',
      type: `${REL}slideMaster`,
      target: 'slideMasters/slideMaster1.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId3',
      type: `${REL}notesMaster`,
      target: 'notesMasters/notesMaster1.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId4',
      type: `${REL}presProps`,
      target: 'presProps.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId5',
      type: `${REL}viewProps`,
      target: 'viewProps.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId6',
      type: `${REL}theme`,
      target: 'theme/theme1.xml',
    });
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rId7',
      type: `${REL}tableStyles`,
      target: 'tableStyles.xml',
    });

    pkg.addRelationship('/ppt/slideMasters/slideMaster1.xml', {
      id: 'rId1',
      type: `${REL}slideLayout`,
      target: '../slideLayouts/slideLayout1.xml',
    });
    pkg.addRelationship('/ppt/slideMasters/slideMaster1.xml', {
      id: 'rId2',
      type: `${REL}theme`,
      target: '../theme/theme1.xml',
    });
    pkg.addRelationship('/ppt/slideLayouts/slideLayout1.xml', {
      id: 'rId1',
      type: `${REL}slideMaster`,
      target: '../slideMasters/slideMaster1.xml',
    });
    pkg.addRelationship('/ppt/notesMasters/notesMaster1.xml', {
      id: 'rId1',
      type: `${REL}theme`,
      target: '../theme/theme1.xml',
    });
    return pkg;
  });
}

function resolveSlideSize(slideSize: BuiltInSlideSize | CustomSlideSize): { readonly cx: number; readonly cy: number } {
  if (typeof slideSize === 'string') {
    if (!Object.hasOwn(SLIDE_SIZES, slideSize)) {
      throw new TypeError(`Unsupported built-in slide size: ${String(slideSize)}`);
    }
    return SLIDE_SIZES[slideSize];
  }
  if (!slideSize || typeof slideSize !== 'object' || Array.isArray(slideSize)) {
    throw new TypeError('Slide size must be a built-in name or a custom width and height');
  }
  return {
    cx: normalizeSlideSizeCoordinate(slideSize.width, 'width'),
    cy: normalizeSlideSizeCoordinate(slideSize.height, 'height'),
  };
}

function normalizeSlideSizeCoordinate(value: unknown, name: 'width' | 'height'): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Custom slide ${name} must be a finite number`);
  }
  const rounded = Math.round(value);
  if (rounded < MIN_SLIDE_SIZE || rounded > MAX_SLIDE_SIZE) {
    throw new RangeError(`Custom slide ${name} must be between 1 and 56 inches`);
  }
  return rounded;
}

function presentationXml(cx: number, cy: number, rtlMode: boolean | undefined): string {
  const rtlAttribute = rtlMode === undefined ? '' : ` rtl="${rtlMode ? '1' : '0'}"`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${rtlAttribute} saveSubsetFonts="1" autoCompressPictures="0"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst/><p:notesMasterIdLst><p:notesMasterId r:id="rId3"/></p:notesMasterIdLst><p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="${cy}" cy="${cx}"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle></p:presentation>`;
}

const CORE_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>@jiayunxie/pptx</dc:creator><cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy><cp:revision>1</cp:revision></cp:coreProperties>`;

const APP_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>@jiayunxie/pptx</Application><PresentationFormat>Custom</PresentationFormat><AppVersion>1.0</AppVersion></Properties>`;

const SLIDE_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:hf sldNum="0" hdr="0" ftr="0" dt="0"/><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`;

const SLIDE_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="DEFAULT"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const NOTES_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf hdr="1" ftr="1" dt="1" sldNum="1"/><p:notesStyle/></p:notesMaster>`;

const PRESENTATION_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;

const VIEW_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView"><p:normalViewPr/><p:slideViewPr/><p:notesTextViewPr/><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;

const TABLE_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="100000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
