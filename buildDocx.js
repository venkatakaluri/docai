const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType } = require('docx');

const CW = 9026;
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
const BORDERS_ALL = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

function run(text, opts = {}) {
  return new TextRun({
    text: String(text || ''),
    bold: opts.bold || false,
    italics: opts.italic || false,
    size: opts.size || 22,
    font: 'Arial',
    color: opts.color || undefined,
  });
}

function para(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before ?? 80, after: opts.after ?? 120 },
    indent: opts.indent ? { left: opts.indent, hanging: opts.hanging || 0 } : undefined,
    border: opts.border || undefined,
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function tcell(text, opts = {}) {
  const { bold = false, shade = null, colW = null, size = 20 } = opts;
  return new TableCell({
    borders: BORDERS_ALL,
    width: colW ? { size: colW, type: WidthType.DXA } : undefined,
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [run(text, { bold, size })],
    })],
  });
}

function answerLines(n) {
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push(new TableRow({
      height: { value: 340, rule: 'exact' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        borders: {
          top:    { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' },
          left:   { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' },
          right:  { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
        },
        children: [new Paragraph({ children: [new TextRun('')] })],
      })],
    }));
  }
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], rows });
}

function makeTable(headers, rows) {
  const n  = Math.max(headers.length, 1);
  const cw = Math.floor(CW / n);
  const widths = headers.map(() => cw);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => tcell(h, { bold: true, shade: 'D6E4F0', colW: widths[i] })),
  });
  const dataRows = (rows || []).map(row =>
    new TableRow({ children: (row || []).map((v, i) => tcell(v, { colW: widths[i] })) })
  );
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...dataRows] });
}

function blockToElements(block) {
  const out = [];
  switch (block.type) {

    case 'title':
      out.push(para(run(block.text, { bold: true, size: 44 }), { align: AlignmentType.CENTER, before: 0, after: 180 }));
      break;

    case 'subtitle':
      out.push(para(run(block.text, { size: 22, color: '555555' }), { align: AlignmentType.CENTER, before: 0, after: 120 }));
      break;

    case 'info_row': {
      const runs = [];
      (block.fields || []).forEach((f, i) => {
        runs.push(run(f + ': ', { bold: true, size: 22 }));
        runs.push(run('_'.repeat(20) + (i < block.fields.length - 1 ? '     ' : ''), { size: 22 }));
      });
      out.push(para(runs, {
        before: 80, after: 200,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, space: 1, color: 'CCCCCC' } }
      }));
      break;
    }

    case 'heading':
      out.push(para(run(block.text, { bold: true, size: 30, color: '1F3864' }), {
        before: 320, after: 120,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, space: 1, color: '2E75B6' } }
      }));
      break;

    case 'subheading':
      out.push(para(run(block.text, { bold: true, size: 24, color: '333333' }), { before: 200, after: 80 }));
      break;

    case 'paragraph':
      out.push(para(run(block.text, { size: 22 }), { before: 60, after: 100 }));
      break;

    case 'note':
      out.push(para(
        [run((block.label || 'Note') + ': ', { bold: true, size: 22 }), run(block.text, { size: 22 })],
        { before: 80, after: 80 }
      ));
      break;

    case 'section':
      out.push(para(run(block.text, { bold: true, size: 24, color: '2E75B6' }), {
        before: 280, after: 100,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, space: 1, color: 'BBBBBB' } }
      }));
      break;

    case 'question': {
      const qNum      = block.number ? `Q${block.number}.  ` : '';
      const marksText = block.marks  ? `[${block.marks} mark${block.marks > 1 ? 's' : ''}]` : '';
      out.push(new Paragraph({
        spacing: { before: 200, after: 60 },
        tabStops: marksText ? [{ type: 'right', position: CW }] : [],
        children: [
          run(qNum, { bold: true, size: 22 }),
          run(block.text, { size: 22 }),
          ...(marksText
            ? [new TextRun({ text: '\t' }), run(marksText, { size: 19, color: '888888', bold: true })]
            : []),
        ],
      }));
      if (block.options && block.options.length) {
        block.options.forEach(opt =>
          out.push(para(run(opt, { size: 22 }), { before: 30, after: 30, indent: 720 }))
        );
      }
      if (block.answer_lines && block.answer_lines > 0) {
        out.push(answerLines(block.answer_lines));
      }
      break;
    }

    case 'bullets':
      (block.items || []).forEach(item =>
        out.push(para(run('\u2022   ' + item, { size: 22 }), { before: 40, after: 40, indent: 540, hanging: 360 }))
      );
      break;

    case 'numbered':
      (block.items || []).forEach((item, i) =>
        out.push(para(run(`${i + 1}.   ${item}`, { size: 22 }), { before: 40, after: 40, indent: 540, hanging: 360 }))
      );
      break;

    case 'table':
      out.push(makeTable(block.headers || [], block.rows || []));
      out.push(new Paragraph({ children: [] }));
      break;

    case 'space':
      out.push(new Paragraph({ spacing: { before: 200, after: 0 }, children: [new TextRun('')] }));
      break;

    default:
      if (block.text) out.push(para(run(block.text, { size: 22 }), { before: 60, after: 100 }));
  }
  return out;
}

async function buildDocx(structure) {
  const children = [];
  for (const block of (structure.blocks || [])) {
    children.push(...blockToElements(block));
  }
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    }],
  });
  return Packer.toBuffer(doc);
}

module.exports = { buildDocx };
