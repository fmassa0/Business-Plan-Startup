/* =====================================================================
   Native Excel chart injection (OOXML)
   ---------------------------------------------------------------------
   ExcelJS does not write charts. We post-process the workbook buffer
   with JSZip: add chart{N}.xml + drawing.xml + relationships, hook them
   to the target sheet. Charts read live data from CE/SP/Indici cells,
   so they recalculate when the user edits inputs in Excel.
   ===================================================================== */
import JSZip from 'jszip';

const NS = {
  c:   'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a:   'http://schemas.openxmlformats.org/drawingml/2006/main',
  r:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
};
const REL_TYPES = {
  chart:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
};
const CT = {
  chart:   'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml'
};

const xmlEsc = (s) => String(s)
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&apos;');

/* ---------- Series builders ---------- */

const numCache = (values, formatCode = 'General') => {
  if (!values || !values.length) return '';
  return `<c:numCache><c:formatCode>${formatCode}</c:formatCode><c:ptCount val="${values.length}"/>${
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${isFinite(v) ? v : 0}</c:v></c:pt>`).join('')
  }</c:numCache>`;
};
const strCache = (values) => {
  if (!values || !values.length) return '';
  return `<c:strCache><c:ptCount val="${values.length}"/>${
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(v)}</c:v></c:pt>`).join('')
  }</c:strCache>`;
};

/* serie: { name, ref, values, color, label, cats, catVals, secondaryAxis } */
function barSeriesXml(serie, idx, opts = {}) {
  const { stack = false, line = false } = opts;
  const colorXml = serie.color
    ? `<c:spPr><a:solidFill><a:srgbClr val="${serie.color}"/></a:solidFill>${line ? `<a:ln w="22000"><a:solidFill><a:srgbClr val="${serie.color}"/></a:solidFill></a:ln>` : ''}</c:spPr>`
    : '';
  const txXml = `<c:tx><c:strRef><c:f>${xmlEsc(serie.nameRef || '""')}</c:f>${strCache([serie.label])}</c:strRef></c:tx>`;
  const catXml = `<c:cat><c:strRef><c:f>${xmlEsc(serie.cats)}</c:f>${strCache(serie.catVals)}</c:strRef></c:cat>`;
  const valXml = `<c:val><c:numRef><c:f>${xmlEsc(serie.ref)}</c:f>${numCache(serie.values)}</c:numRef></c:val>`;
  if (line) {
    return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${txXml}${colorXml}` +
      `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${serie.color || '0A1F44'}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>` +
      `<c:smooth val="0"/>${catXml}${valXml}</c:ser>`;
  }
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>${txXml}${colorXml}<c:invertIfNegative val="0"/>${catXml}${valXml}</c:ser>`;
}

/* ---------- Chart builders ---------- */

/** Combo chart: bar + line (e.g. Ricavi/EBITDA bars + Utile line). */
function chartXml_BarLine({ title, barSeries, lineSeries, axId1 = 11111, axId2 = 22222, valFormat = '#,##0' }) {
  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:r><a:rPr lang="it-IT" sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill></a:rPr><a:t>${xmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : '<c:autoTitleDeleted val="1"/>';

  const barXml = barSeries && barSeries.length
    ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${
        barSeries.map((s, i) => barSeriesXml(s, i)).join('')
      }<c:gapWidth val="80"/><c:axId val="${axId1}"/><c:axId val="${axId2}"/></c:barChart>`
    : '';
  const lineXml = lineSeries && lineSeries.length
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${
        lineSeries.map((s, i) => barSeriesXml(s, (barSeries ? barSeries.length : 0) + i, { line: true })).join('')
      }<c:marker val="1"/><c:axId val="${axId1}"/><c:axId val="${axId2}"/></c:lineChart>`
    : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      ${barXml}
      ${lineXml}
      <c:catAx>
        <c:axId val="${axId1}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axId2}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${axId2}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines><c:spPr><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="EEF1F5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:numFmt formatCode="${xmlEsc(valFormat)}" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axId1}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;
}

/** Stacked bar + line (e.g. SP structure) */
function chartXml_StackedBarLine({ title, stackedSeries, lineSeries, axId1 = 33333, axId2 = 44444, valFormat = '#,##0' }) {
  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="it-IT" sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill></a:rPr><a:t>${xmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : '<c:autoTitleDeleted val="1"/>';

  const stackedBarXml = stackedSeries && stackedSeries.length
    ? `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>${
        stackedSeries.map((s, i) => barSeriesXml(s, i)).join('')
      }<c:gapWidth val="80"/><c:overlap val="100"/><c:axId val="${axId1}"/><c:axId val="${axId2}"/></c:barChart>`
    : '';
  const lineXml = lineSeries && lineSeries.length
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${
        lineSeries.map((s, i) => barSeriesXml(s, (stackedSeries ? stackedSeries.length : 0) + i, { line: true })).join('')
      }<c:marker val="1"/><c:axId val="${axId1}"/><c:axId val="${axId2}"/></c:lineChart>`
    : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      ${stackedBarXml}
      ${lineXml}
      <c:catAx>
        <c:axId val="${axId1}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axId2}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${axId2}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="EEF1F5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:numFmt formatCode="${xmlEsc(valFormat)}" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axId1}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;
}

/** Dual-axis line chart (e.g. DSCR + PFN/EBITDA on different scales) */
function chartXml_DualAxisLine({ title, primarySeries, secondarySeries, primaryLabel, secondaryLabel, axIdC = 55555, axIdV1 = 66666, axIdV2 = 77777, axIdC2 = 88888, primaryFmt = '0.00', secondaryFmt = '0.00' }) {
  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="it-IT" sz="1100" b="1"><a:solidFill><a:srgbClr val="0A1F44"/></a:solidFill></a:rPr><a:t>${xmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : '<c:autoTitleDeleted val="1"/>';

  const primaryLineXml = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${
    primarySeries.map((s, i) => barSeriesXml(s, i, { line: true })).join('')
  }<c:marker val="1"/><c:axId val="${axIdC}"/><c:axId val="${axIdV1}"/></c:lineChart>`;

  const secondaryLineXml = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${
    secondarySeries.map((s, i) => barSeriesXml(s, primarySeries.length + i, { line: true })).join('')
  }<c:marker val="1"/><c:axId val="${axIdC2}"/><c:axId val="${axIdV2}"/></c:lineChart>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      ${primaryLineXml}
      ${secondaryLineXml}
      <c:catAx>
        <c:axId val="${axIdC}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axIdV1}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${axIdV1}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="EEF1F5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="00B0A0"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="it-IT" sz="900"><a:solidFill><a:srgbClr val="00B0A0"/></a:solidFill></a:rPr><a:t>${xmlEsc(primaryLabel || 'Primary')}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
        <c:numFmt formatCode="${xmlEsc(primaryFmt)}" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axIdC}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="${axIdV2}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="r"/>
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="EF4444"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="it-IT" sz="900"><a:solidFill><a:srgbClr val="EF4444"/></a:solidFill></a:rPr><a:t>${xmlEsc(secondaryLabel || 'Secondary')}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
        <c:numFmt formatCode="${xmlEsc(secondaryFmt)}" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="it-IT"/></a:p></c:txPr>
        <c:crossAx val="${axIdC2}"/>
        <c:crosses val="max"/>
        <c:crossBetween val="between"/>
      </c:valAx>
      <c:catAx>
        <c:axId val="${axIdC2}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="1"/>
        <c:axPos val="b"/>
        <c:crossAx val="${axIdV2}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;
}

/* ---------- Drawing & rels ---------- */

function drawingXmlForCharts(chartCount, positions) {
  const anchors = positions.map((pos, i) => `
    <xdr:twoCellAnchor editAs="oneCell">
      <xdr:from><xdr:col>${pos.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pos.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>${pos.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pos.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
      <xdr:graphicFrame macro="">
        <xdr:nvGraphicFramePr>
          <xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/>
          <xdr:cNvGraphicFramePr/>
        </xdr:nvGraphicFramePr>
        <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
        <a:graphic>
          <a:graphicData uri="${NS.c}">
            <c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="rId${i + 1}"/>
          </a:graphicData>
        </a:graphic>
      </xdr:graphicFrame>
      <xdr:clientData/>
    </xdr:twoCellAnchor>`).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${anchors}
</xdr:wsDr>`;
}

function drawingRelsXml(chartCount, drawingNumber) {
  // chart files indexed starting from a base; we'll pass actual chart numbers
  const rels = [];
  for (let i = 0; i < chartCount; i++) {
    const chartNum = drawingNumber === 1 ? i + 1 : (i + 1); // assuming charts dedicated to this drawing
    // Caller must provide chart numbers; here we use 1..N
  }
  return null; // see relsXmlFromList
}

function relsXmlFromList(items) {
  const body = items.map((it) => `<Relationship Id="${it.id}" Type="${it.type}" Target="${it.target}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

/* ---------- Workbook injection ---------- */

/**
 * Inject native charts into a workbook buffer.
 * @param {ArrayBuffer|Uint8Array|Buffer} workbookBuffer  Output of ExcelJS writeBuffer
 * @param {object} opts
 * @param {string} opts.targetSheetName  Sheet that will host the charts (must exist)
 * @param {Array<{xml:string, position:{fromCol,fromRow,toCol,toRow}}>} opts.charts
 * @returns {Promise<Blob>}
 */
export async function injectNativeCharts(workbookBuffer, { targetSheetName, charts }) {
  const zip = await JSZip.loadAsync(workbookBuffer);

  // 1) Resolve target worksheet path via workbook.xml + workbook rels
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbRelsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');

  const sheetMatch = new RegExp(`<sheet[^>]*name="${targetSheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*r:id="(rId\\d+)"`, 'i').exec(wbXml);
  if (!sheetMatch) throw new Error(`Sheet "${targetSheetName}" non trovato nel workbook`);
  const sheetRid = sheetMatch[1];

  const relMatch = new RegExp(`<Relationship[^>]*Id="${sheetRid}"[^>]*Target="([^"]+)"`, 'i').exec(wbRelsXml);
  if (!relMatch) throw new Error(`Relationship per ${sheetRid} non trovato`);
  let sheetTarget = relMatch[1]; // e.g. "worksheets/sheet7.xml"
  if (!sheetTarget.startsWith('xl/')) sheetTarget = 'xl/' + sheetTarget;
  const sheetFileName = sheetTarget.split('/').pop(); // sheet7.xml
  const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;

  // 2) Find next free chart/drawing numbers
  const existingCharts = Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f));
  const existingDrawings = Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f));
  const nextChartIdx = existingCharts.length + 1;
  const nextDrawingIdx = existingDrawings.length + 1;

  // 3) Write chart{N}.xml
  const chartNumbers = charts.map((_, i) => nextChartIdx + i);
  charts.forEach((c, i) => {
    zip.file(`xl/charts/chart${chartNumbers[i]}.xml`, c.xml);
  });

  // 4) Write drawing{N}.xml + drawing rels
  const positions = charts.map((c) => c.position);
  const drawingXml = drawingXmlForCharts(charts.length, positions);
  zip.file(`xl/drawings/drawing${nextDrawingIdx}.xml`, drawingXml);

  const drawingRels = relsXmlFromList(
    chartNumbers.map((n, i) => ({
      id: `rId${i + 1}`,
      type: REL_TYPES.chart,
      target: `../charts/chart${n}.xml`
    }))
  );
  zip.file(`xl/drawings/_rels/drawing${nextDrawingIdx}.xml.rels`, drawingRels);

  // 5) Update worksheet rels: add drawing rel
  let sheetRelsContent;
  const existingSheetRels = zip.file(sheetRelsPath);
  if (existingSheetRels) {
    sheetRelsContent = await existingSheetRels.async('string');
  } else {
    sheetRelsContent = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  }
  // determine next rId number
  const existingRids = [...sheetRelsContent.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
  const nextRid = (existingRids.length ? Math.max(...existingRids) : 0) + 1;
  const drawingRelLine = `<Relationship Id="rId${nextRid}" Type="${REL_TYPES.drawing}" Target="../drawings/drawing${nextDrawingIdx}.xml"/>`;
  const updatedSheetRels = sheetRelsContent.replace(/<\/Relationships>/, drawingRelLine + '</Relationships>');
  zip.file(sheetRelsPath, updatedSheetRels);

  // 6) Update worksheet XML to include <drawing r:id="rIdN"/>
  // Schema order: drawing must come AFTER pageSetup/headerFooter etc., BEFORE extLst.
  let sheetXmlContent = await zip.file(sheetTarget).async('string');
  const drawingTag = `<drawing r:id="rId${nextRid}"/>`;
  if (/<drawing[\s/]/.test(sheetXmlContent)) {
    sheetXmlContent = sheetXmlContent.replace(/<drawing[^/]*\/>/, drawingTag);
  } else if (/<extLst>/i.test(sheetXmlContent)) {
    sheetXmlContent = sheetXmlContent.replace(/<extLst>/i, drawingTag + '<extLst>');
  } else if (/<legacyDrawing[\s/]/i.test(sheetXmlContent)) {
    sheetXmlContent = sheetXmlContent.replace(/<legacyDrawing/i, drawingTag + '<legacyDrawing');
  } else {
    sheetXmlContent = sheetXmlContent.replace(/<\/worksheet>/, drawingTag + '</worksheet>');
  }
  zip.file(sheetTarget, sheetXmlContent);

  // 7) Update [Content_Types].xml
  let ctXml = await zip.file('[Content_Types].xml').async('string');
  chartNumbers.forEach((n) => {
    const partName = `/xl/charts/chart${n}.xml`;
    if (!ctXml.includes(`PartName="${partName}"`)) {
      ctXml = ctXml.replace(/<\/Types>/, `<Override PartName="${partName}" ContentType="${CT.chart}"/></Types>`);
    }
  });
  const drawingPartName = `/xl/drawings/drawing${nextDrawingIdx}.xml`;
  if (!ctXml.includes(`PartName="${drawingPartName}"`)) {
    ctXml = ctXml.replace(/<\/Types>/, `<Override PartName="${drawingPartName}" ContentType="${CT.drawing}"/></Types>`);
  }
  zip.file('[Content_Types].xml', ctXml);

  // 8) Re-zip
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    compression: 'DEFLATE'
  });
}

/* ---------- Public chart-XML factories ---------- */

export function makeChartXml_RicaviEbitdaUtile({ ceSheetName, headerRow, rowRic, rowEbitda, rowUtile, yearLabels, ricavi, ebitda, utile, palette }) {
  const sheet = `'${ceSheetName}'`;
  const cats = `${sheet}!$B$${headerRow}:$F$${headerRow}`;
  return chartXml_BarLine({
    title: 'Ricavi vs EBITDA vs Utile Netto',
    barSeries: [
      { label: 'Ricavi',  nameRef: `${sheet}!$A$${rowRic}`,   ref: `${sheet}!$B$${rowRic}:$F$${rowRic}`,     values: ricavi, color: palette.navy,  cats, catVals: yearLabels },
      { label: 'EBITDA',  nameRef: `${sheet}!$A$${rowEbitda}`, ref: `${sheet}!$B$${rowEbitda}:$F$${rowEbitda}`, values: ebitda, color: palette.teal, cats, catVals: yearLabels }
    ],
    lineSeries: [
      { label: 'Utile Netto', nameRef: `${sheet}!$A$${rowUtile}`, ref: `${sheet}!$B$${rowUtile}:$F$${rowUtile}`, values: utile, color: palette.amber, cats, catVals: yearLabels }
    ]
  });
}

export function makeChartXml_FCF({ indSheetName, headerRow, rowFCF, rowFCFcum, yearLabels, fcf, fcfcum, palette }) {
  const sheet = `'${indSheetName}'`;
  const cats = `${sheet}!$B$${headerRow}:$F$${headerRow}`;
  return chartXml_BarLine({
    title: 'Free Cash Flow & FCF Cumulato',
    axId1: 12121, axId2: 23232,
    barSeries: [
      { label: 'FCF', nameRef: `${sheet}!$A$${rowFCF}`, ref: `${sheet}!$B$${rowFCF}:$F$${rowFCF}`, values: fcf, color: palette.teal, cats, catVals: yearLabels }
    ],
    lineSeries: [
      { label: 'FCF cumulato', nameRef: `${sheet}!$A$${rowFCFcum}`, ref: `${sheet}!$B$${rowFCFcum}:$F$${rowFCFcum}`, values: fcfcum, color: palette.navy, cats, catVals: yearLabels }
    ]
  });
}

export function makeChartXml_StatoPatrimoniale({ spSheetName, headerRow, rowPN, rowDebFin, rowDebTot, rowCassa, yearLabels, pn, debFin, debTot, cassa, palette }) {
  const sheet = `'${spSheetName}'`;
  // SP usa colonne C..G per anni 1..5
  const cats = `${sheet}!$C$${headerRow}:$G$${headerRow}`;
  const rng = (row) => `${sheet}!$C$${row}:$G$${row}`;
  return chartXml_StackedBarLine({
    title: 'Struttura Stato Patrimoniale',
    axId1: 34343, axId2: 45454,
    stackedSeries: [
      { label: 'Patrimonio Netto',          nameRef: `${sheet}!$A$${rowPN}`,     ref: rng(rowPN),     values: pn,     color: palette.navy,  cats, catVals: yearLabels },
      { label: 'Debiti finanziari',         nameRef: `${sheet}!$A$${rowDebFin}`, ref: rng(rowDebFin), values: debFin, color: palette.red,   cats, catVals: yearLabels },
      { label: 'Altri debiti (commerc.+IVA+trib.)', nameRef: `${sheet}!$A$${rowDebTot}`, ref: rng(rowDebTot), values: debTot, color: palette.amber, cats, catVals: yearLabels }
    ],
    lineSeries: [
      { label: 'Cassa', nameRef: `${sheet}!$A$${rowCassa}`, ref: rng(rowCassa), values: cassa, color: palette.teal, cats, catVals: yearLabels }
    ]
  });
}

export function makeChartXml_DSCRpfn({ indSheetName, headerRow, rowDSCR, rowPFNebitda, yearLabels, dscr, pfnEbitda, palette }) {
  const sheet = `'${indSheetName}'`;
  const cats = `${sheet}!$B$${headerRow}:$F$${headerRow}`;
  return chartXml_DualAxisLine({
    title: 'Indici di Sostenibilità',
    primaryLabel: 'DSCR (x)',
    secondaryLabel: 'PFN/EBITDA (x)',
    primarySeries: [
      { label: 'DSCR', nameRef: `${sheet}!$A$${rowDSCR}`, ref: `${sheet}!$B$${rowDSCR}:$F$${rowDSCR}`, values: dscr, color: palette.teal, cats, catVals: yearLabels }
    ],
    secondarySeries: [
      { label: 'PFN/EBITDA', nameRef: `${sheet}!$A$${rowPFNebitda}`, ref: `${sheet}!$B$${rowPFNebitda}:$F$${rowPFNebitda}`, values: pfnEbitda, color: palette.red, cats, catVals: yearLabels }
    ]
  });
}
