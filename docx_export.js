const {
  AlignmentType,
  Document,
  FootnoteReferenceRun,
  HeadingLevel,
  HighlightColor,
  LevelFormat,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType
} = require('docx');

const ALIGNMENT_MAP = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED
};

const HEADING_MAP = {
  heading1: HeadingLevel.HEADING_1,
  heading2: HeadingLevel.HEADING_2,
  heading3: HeadingLevel.HEADING_3,
  heading4: HeadingLevel.HEADING_4
};

const HIGHLIGHT_HEX_MAP = [
  { key: HighlightColor.YELLOW, hex: 'FFFF00' },
  { key: HighlightColor.GREEN, hex: '00FF00' },
  { key: HighlightColor.CYAN, hex: '00FFFF' },
  { key: HighlightColor.MAGENTA, hex: 'FF00FF' },
  { key: HighlightColor.BLUE, hex: '0000FF' },
  { key: HighlightColor.RED, hex: 'FF0000' },
  { key: HighlightColor.DARK_YELLOW, hex: '808000' },
  { key: HighlightColor.DARK_GREEN, hex: '008000' },
  { key: HighlightColor.DARK_CYAN, hex: '008080' },
  { key: HighlightColor.DARK_MAGENTA, hex: '800080' },
  { key: HighlightColor.DARK_BLUE, hex: '000080' },
  { key: HighlightColor.DARK_RED, hex: '800000' },
  { key: HighlightColor.DARK_GRAY, hex: '404040' },
  { key: HighlightColor.LIGHT_GRAY, hex: 'C0C0C0' },
  { key: HighlightColor.BLACK, hex: '000000' },
  { key: HighlightColor.WHITE, hex: 'FFFFFF' }
];

function normalizeHexColor(value) {
  if (typeof value !== 'string') return undefined;
  const raw = value.trim().replace(/^#/, '');
  if (/^[0-9a-fA-F]{6}$/.test(raw)) return raw.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(raw)) {
    return raw
      .split('')
      .map(function(ch) { return ch + ch; })
      .join('')
      .toUpperCase();
  }
  return undefined;
}

function parseHexRgb(hex) {
  const normalized = normalizeHexColor(hex);
  if (!normalized) return null;
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16)
  };
}

function mapHighlightColor(hex) {
  const rgb = parseHexRgb(hex);
  if (!rgb) return undefined;

  let best = null;
  let bestDist = Number.POSITIVE_INFINITY;
  for (let i = 0; i < HIGHLIGHT_HEX_MAP.length; i += 1) {
    const candidate = HIGHLIGHT_HEX_MAP[i];
    const crgb = parseHexRgb(candidate.hex);
    if (!crgb) continue;
    const dr = rgb.r - crgb.r;
    const dg = rgb.g - crgb.g;
    const db = rgb.b - crgb.b;
    const dist = dr * dr + dg * dg + db * db;
    if (dist < bestDist) {
      bestDist = dist;
      best = candidate.key;
    }
  }
  return best || undefined;
}

function clampNumber(value, min, max) {
  if (typeof value !== 'number' || Number.isNaN(value)) return min;
  return Math.max(min, Math.min(max, value));
}

function toHalfPoints(sizePt) {
  if (typeof sizePt !== 'number' || Number.isNaN(sizePt)) return undefined;
  const clamped = clampNumber(sizePt, 1, 400);
  return Math.round(clamped * 2);
}

function toTwips(points) {
  if (typeof points !== 'number' || Number.isNaN(points)) return undefined;
  return Math.round(points * 20);
}

function normalizeRunText(raw) {
  if (typeof raw !== 'string') return '';
  return raw.replace(/\u200B/g, '');
}

function pushTextRunsFromText(outRuns, baseRun, rawText) {
  const normalized = normalizeRunText(rawText);
  if (!normalized) return;

  const pieces = normalized.split('\n');
  for (let i = 0; i < pieces.length; i += 1) {
    if (i > 0) {
      outRuns.push(new TextRun({ break: 1 }));
    }
    const piece = pieces[i];
    if (!piece) continue;
    outRuns.push(new TextRun(Object.assign({}, baseRun, { text: piece })));
  }
}

function buildRuns(runModels) {
  const out = [];
  if (!Array.isArray(runModels)) return out;

  for (let i = 0; i < runModels.length; i += 1) {
    const run = runModels[i] || {};
    if (run.footnoteRef) {
      const noteId = Number(run.footnoteRef);
      if (Number.isFinite(noteId) && noteId > 0) {
        out.push(new FootnoteReferenceRun(noteId));
      }
      continue;
    }
    const runBase = {};

    if (run.bold) runBase.bold = true;
    if (run.italic) runBase.italics = true;
    if (run.underline) runBase.underline = { type: UnderlineType.SINGLE };
    if (run.strike) runBase.strike = true;
    if (run.subscript) runBase.subScript = true;
    if (run.superscript) runBase.superScript = true;

    const color = normalizeHexColor(run.color);
    if (color) runBase.color = color;

    const highlight = mapHighlightColor(run.highlight);
    if (highlight) runBase.highlight = highlight;

    if (typeof run.font === 'string' && run.font.trim()) {
      runBase.font = run.font.trim();
    }

    const halfPoints = toHalfPoints(run.sizePt);
    if (halfPoints) runBase.size = halfPoints;

    if (run.breaks && run.breaks > 0) {
      const breaks = Math.floor(run.breaks);
      for (let b = 0; b < breaks; b += 1) {
        out.push(new TextRun({ break: 1 }));
      }
    }

    pushTextRunsFromText(out, runBase, run.text || '');
  }

  return out;
}

function buildFootnoteParagraphs(text) {
  const raw = String(text || '').replace(/\r/g, '');
  const lines = raw.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i += 1) {
    out.push(new Paragraph({ children: [new TextRun(lines[i] || '')] }));
  }
  if (out.length === 0) out.push(new Paragraph({ children: [new TextRun('')] }));
  return out;
}

function buildFootnotesMap(payloadFootnotes) {
  const map = {};
  const notes = Array.isArray(payloadFootnotes) ? payloadFootnotes : [];
  for (let i = 0; i < notes.length; i += 1) {
    const note = notes[i] || {};
    const id = Number(note.id);
    if (!Number.isFinite(id) || id <= 0) continue;
    map[String(id)] = {
      children: buildFootnoteParagraphs(note.text || '')
    };
  }
  return map;
}

function isParagraphLike(block) {
  return !!(block && block.type === 'paragraph');
}

function buildParagraph(block) {
  const options = {};
  const runs = buildRuns(block.runs || []);
  options.children = runs.length > 0 ? runs : [new TextRun('')];

  if (typeof block.heading === 'string') {
    const heading = HEADING_MAP[block.heading.toLowerCase()];
    if (heading) options.heading = heading;
  }

  if (typeof block.alignment === 'string') {
    const alignment = ALIGNMENT_MAP[block.alignment.toLowerCase()];
    if (alignment) options.alignment = alignment;
  }

  const spacing = {};
  if (typeof block.lineSpacing === 'number' && !Number.isNaN(block.lineSpacing)) {
    const multiple = clampNumber(block.lineSpacing, 0.5, 4);
    spacing.line = Math.round(multiple * 240);
  }
  if (Object.keys(spacing).length > 0) options.spacing = spacing;

  const indent = {};
  if (typeof block.indentLeftPt === 'number' && !Number.isNaN(block.indentLeftPt)) {
    indent.left = Math.max(0, toTwips(block.indentLeftPt) || 0);
  }
  if (typeof block.indentFirstPt === 'number' && !Number.isNaN(block.indentFirstPt)) {
    indent.firstLine = toTwips(block.indentFirstPt) || 0;
  }
  if (Object.keys(indent).length > 0) options.indent = indent;

  if (block.list && typeof block.list === 'object') {
    const listType = String(block.list.type || '').toLowerCase();
    const level = clampNumber(Number(block.list.level || 0), 0, 8);
    if (listType === 'bullet' || listType === 'number') {
      options.numbering = {
        reference: listType === 'bullet' ? 'parallel-bullet' : 'parallel-number',
        level: Math.floor(level)
      };
    }
  }

  return new Paragraph(options);
}

function buildCellParagraphs(cell) {
  const blocks = Array.isArray(cell && cell.blocks) ? cell.blocks : [];
  const paragraphs = [];
  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    if (isParagraphLike(block)) {
      paragraphs.push(buildParagraph(block));
    }
  }
  if (paragraphs.length === 0) paragraphs.push(new Paragraph(''));
  return paragraphs;
}

function buildTable(block) {
  const rowModels = Array.isArray(block.rows) ? block.rows : [];
  const rows = [];

  for (let i = 0; i < rowModels.length; i += 1) {
    const rowModel = rowModels[i];
    const cellModels = Array.isArray(rowModel && rowModel.cells) ? rowModel.cells : [];
    const cells = [];
    for (let c = 0; c < cellModels.length; c += 1) {
      cells.push(new TableCell({
        children: buildCellParagraphs(cellModels[c]),
        width: { size: 100 / Math.max(1, cellModels.length), type: WidthType.PERCENTAGE }
      }));
    }
    if (cells.length > 0) rows.push(new TableRow({ children: cells }));
  }

  if (rows.length === 0) {
    rows.push(new TableRow({
      children: [new TableCell({ children: [new Paragraph('')] })]
    }));
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows
  });
}

function hasTextInRuns(runs) {
  if (!Array.isArray(runs)) return false;
  for (let i = 0; i < runs.length; i += 1) {
    const run = runs[i];
    if (!run) continue;
    if (typeof run.text === 'string' && run.text.replace(/\u200B/g, '').trim()) return true;
  }
  return false;
}

function sectionHasContent(section) {
  if (!section || !Array.isArray(section.blocks)) return false;
  for (let i = 0; i < section.blocks.length; i += 1) {
    const block = section.blocks[i];
    if (!block) continue;
    if (block.type === 'paragraph' && hasTextInRuns(block.runs || [])) return true;
    if (block.type === 'table') return true;
  }
  return false;
}

function normalizeSections(rawSections) {
  if (!Array.isArray(rawSections)) return [];
  const ordered = [];
  for (let i = 0; i < rawSections.length; i += 1) {
    const section = rawSections[i];
    if (!section || !Array.isArray(section.blocks)) continue;
    ordered.push(section);
  }
  return ordered.filter(sectionHasContent);
}

function buildDocumentChildren(sections) {
  const children = [];
  for (let s = 0; s < sections.length; s += 1) {
    const section = sections[s];
    if (s > 0) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    for (let i = 0; i < section.blocks.length; i += 1) {
      const block = section.blocks[i];
      if (!block) continue;
      if (block.type === 'paragraph') {
        children.push(buildParagraph(block));
      } else if (block.type === 'table') {
        children.push(buildTable(block));
      }
    }
  }
  if (children.length === 0) children.push(new Paragraph(''));
  return children;
}

function buildNumberingConfig() {
  const bulletSymbols = ['•', 'o', '▪', '•', 'o', '▪', '•', 'o', '▪'];
  const bulletLevels = [];
  const numberLevels = [];

  for (let level = 0; level <= 8; level += 1) {
    bulletLevels.push({
      level,
      format: LevelFormat.BULLET,
      text: bulletSymbols[level] || '•',
      alignment: AlignmentType.LEFT,
      style: {
        paragraph: {
          indent: {
            left: 720 + (level * 360),
            hanging: 260
          }
        }
      }
    });

    numberLevels.push({
      level,
      format: LevelFormat.DECIMAL,
      text: `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      style: {
        paragraph: {
          indent: {
            left: 720 + (level * 360),
            hanging: 260
          }
        }
      }
    });
  }

  return [
    { reference: 'parallel-bullet', levels: bulletLevels },
    { reference: 'parallel-number', levels: numberLevels }
  ];
}

async function buildDocxBufferFromModel(payload) {
  const sections = normalizeSections(payload && payload.sections);
  if (sections.length === 0) {
    throw new Error('No section content available to export.');
  }
  const footnotes = buildFootnotesMap(payload && payload.footnotes);
  const hasFootnotes = Object.keys(footnotes).length > 0;

  const doc = new Document({
    numbering: {
      config: buildNumberingConfig()
    },
    footnotes: hasFootnotes ? footnotes : undefined,
    sections: [
      {
        properties: {},
        children: buildDocumentChildren(sections)
      }
    ]
  });

  return Packer.toBuffer(doc);
}

module.exports = {
  buildDocxBufferFromModel
};
