// exam.js — dựng OOXML cho các khối phục vụ soạn đề thi:
// đầu đề thi, câu hỏi theo cấu trúc đề THPT, bảng biến thiên/xét dấu, bảng đáp án.
// Tất cả đều đi qua wrapBody() của symbols.js để thành gói Flat OPC hoàn chỉnh.

// Bề ngang vùng soạn thảo của A4 lề 2cm, tính bằng twip (1/20 pt).
const TWIP_PAGE = 9070;

function esc(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function run(content, opts = {}) {
  const rpr =
    (opts.bold ? '<w:b/>' : '') +
    (opts.italic ? '<w:i/>' : '') +
    (opts.size ? `<w:sz w:val="${opts.size * 2}"/>` : '');
  const props = rpr ? `<w:rPr>${rpr}</w:rPr>` : '';
  return `<w:r>${props}<w:t xml:space="preserve">${esc(content)}</w:t></w:r>`;
}

// Thứ tự các con của w:pPr phải đúng theo schema OOXML (spacing → ind → jc),
// sai thứ tự thì Word từ chối cả gói với thông báo "problem with its contents".
function para(runsXml, opts = {}) {
  const ppr =
    (opts.spaceAfter !== undefined ? `<w:spacing w:after="${opts.spaceAfter}"/>` : '') +
    (opts.ind ? `<w:ind w:left="${opts.ind}"/>` : '') +
    (opts.jc ? `<w:jc w:val="${opts.jc}"/>` : '');
  return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runsXml}</w:p>`;
}

const textPara = (content, opts = {}) => para(run(content, opts), opts);

const BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];

function borderXml(tag, sides) {
  const inner = Object.entries(sides)
    .map(([side, val]) => `<w:${side} w:val="${val}" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('');
  return `<w:${tag}>${inner}</w:${tag}>`;
}

function tc(width, contentXml, opts = {}) {
  const borders = opts.borders ? borderXml('tcBorders', opts.borders) : '';
  return (
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${borders}` +
    `<w:vAlign w:val="${opts.vAlign || 'center'}"/></w:tcPr>` +
    `${contentXml || '<w:p/>'}</w:tc>`
  );
}

function tr(cellsXml, height) {
  const props = height ? `<w:trPr><w:trHeight w:val="${height}"/></w:trPr>` : '';
  return `<w:tr>${props}${cellsXml}</w:tr>`;
}

// Word đòi phải có một đoạn văn ngay sau bảng, nếu không tài liệu sẽ hỏng.
function tbl(widths, rowsXml, bordered) {
  const sides = {};
  BORDER_SIDES.forEach((side) => {
    sides[side] = bordered ? 'single' : 'none';
  });
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' +
    borderXml('tblBorders', sides) +
    '<w:tblLayout w:type="fixed"/></w:tblPr>' +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    `${rowsXml}</w:tbl><w:p/>`
  );
}

// ---------------------------------------------------------------- đầu đề thi

function examHeader(f) {
  const half = Math.floor(TWIP_PAGE / 2);
  const centred = { jc: 'center', spaceAfter: 0 };
  const left =
    textPara(f.so || 'SỞ GD&ĐT .....................', centred) +
    textPara(f.truong || 'TRƯỜNG .....................', { ...centred, bold: true }) +
    textPara(`(Đề thi gồm ${f.soTrang || '...'} trang)`, { ...centred, italic: true });
  const right =
    textPara(f.kyThi || 'ĐỀ KIỂM TRA CUỐI HỌC KỲ I', { ...centred, bold: true }) +
    textPara(`Môn: ${f.mon || 'Toán'}${f.khoi ? ` — Lớp ${f.khoi}` : ''}`, {
      ...centred,
      bold: true,
    }) +
    textPara(`Thời gian làm bài: ${f.thoiGian || '90'} phút`, { ...centred, italic: true });

  const table = tbl(
    [half, half],
    tr(tc(half, left, { vAlign: 'top' }) + tc(half, right, { vAlign: 'top' })),
    false
  );

  return wrapBody(
    table +
      textPara(`Mã đề: ${f.maDe || '...'}`, { jc: 'right', bold: true }) +
      textPara(
        'Họ và tên thí sinh: ..................................................  ' +
          'Số báo danh: ....................'
      ) +
      '<w:p/>'
  );
}

// ---------------------------------------------------------------- câu hỏi

const SECTIONS = {
  1:
    'PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu ..., ' +
    'mỗi câu chỉ chọn một phương án.',
  2:
    'PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu ... đến câu ..., ' +
    'trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai.',
  3: 'PHẦN III. Câu trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu ... đến câu ....',
};

function sectionHeading(which) {
  return wrapBody(textPara(SECTIONS[which], { bold: true }));
}

const MC_LABELS = ['A.', 'B.', 'C.', 'D.'];
const TF_LABELS = ['a)', 'b)', 'c)', 'd)'];

function multipleChoice(no, cols) {
  const stem = textPara(`Câu ${no}: `, { bold: true, spaceAfter: 0 });
  if (cols === 1) {
    return (
      stem + MC_LABELS.map((l) => textPara(`${l} `, { ind: 284, spaceAfter: 0 })).join('') + '<w:p/>'
    );
  }
  const width = Math.floor(TWIP_PAGE / cols);
  const cells = MC_LABELS.map((l) =>
    tc(width, textPara(`${l} `, { spaceAfter: 0 }), { vAlign: 'top' })
  );
  const rows = [];
  for (let i = 0; i < cells.length; i += cols) rows.push(tr(cells.slice(i, i + cols).join('')));
  return stem + tbl(new Array(cols).fill(width), rows.join(''), false);
}

function trueFalse(no) {
  return (
    textPara(`Câu ${no}: `, { bold: true, spaceAfter: 0 }) +
    TF_LABELS.map((l) => textPara(`${l} `, { ind: 284, spaceAfter: 0 })).join('') +
    '<w:p/>'
  );
}

function shortAnswer(no) {
  return textPara(`Câu ${no}: `, { bold: true, spaceAfter: 0 }) + '<w:p/>';
}

const QUESTION_BUILDERS = { mc: multipleChoice, tf: trueFalse, sa: shortAnswer };

function questionBlock(kind, start, count, cols) {
  const build = QUESTION_BUILDERS[kind];
  let out = '';
  for (let i = 0; i < count; i += 1) out += build(start + i, cols);
  return wrapBody(out);
}

// ---------------------------------------------------------------- bảng đáp án

// Chấp nhận "1A 2B 3C", "1.A, 2.B", "ABCD..." hoặc đáp án đúng/sai kiểu "1 ĐSSĐ".
function parseAnswerKey(raw, start) {
  const src = String(raw || '').trim();
  if (!src) return [];
  if (/\d/.test(src)) {
    const out = [];
    const re = /(\d+)\s*[.:)\-]?\s*([A-Da-dĐSđs]+)/g;
    let m = re.exec(src);
    while (m) {
      out.push({ no: Number(m[1]), ans: m[2].toUpperCase() });
      m = re.exec(src);
    }
    return out;
  }
  return (src.match(/[A-Da-d]/g) || []).map((ans, i) => ({
    no: start + i,
    ans: ans.toUpperCase(),
  }));
}

function answerKeyTable(entries, perRow) {
  const labelW = 760;
  const cellW = Math.floor((TWIP_PAGE - labelW) / perRow);
  const widths = [labelW, ...new Array(perRow).fill(cellW)];
  let rows = '';
  for (let i = 0; i < entries.length; i += perRow) {
    const band = entries.slice(i, i + perRow);
    const blanks = new Array(perRow - band.length).fill(tc(cellW, '')).join('');
    const heads = band.map((e) => tc(cellW, textPara(String(e.no), { jc: 'center', bold: true })));
    const cells = band.map((e) => tc(cellW, textPara(e.ans, { jc: 'center' })));
    rows += tr(tc(labelW, textPara('Câu', { jc: 'center', bold: true })) + heads.join('') + blanks);
    rows += tr(
      tc(labelW, textPara('Đáp án', { jc: 'center', bold: true })) + cells.join('') + blanks
    );
  }
  return wrapBody(textPara('BẢNG ĐÁP ÁN', { jc: 'center', bold: true }) + tbl(widths, rows, true));
}

// ------------------------------------------------- bảng biến thiên / xét dấu

const ARROW = { '+': '↗', '-': '↘' };

// cfg: { mode:'bbt'|'sign', varName, fnName, derivName, nodes:[...], signs:[...], marks:[...] }
// nodes gồm cả hai đầu mút; signs có đúng nodes.length - 1 phần tử, marks bằng nodes.
function variationTable(cfg) {
  const { nodes, signs, marks } = cfg;
  const labelW = 1000;
  const nodeW = 620;
  const gapW = Math.max(
    700,
    Math.floor((TWIP_PAGE - labelW - nodeW * nodes.length) / signs.length)
  );

  const widths = [labelW];
  const kinds = []; // 'node' | 'gap' — mô tả từng ô thân bảng
  nodes.forEach((_, i) => {
    widths.push(nodeW);
    kinds.push('node');
    if (i < signs.length) {
      widths.push(gapW);
      kinds.push('gap');
    }
  });

  // openSides: bỏ các đường kẻ dọc bên trong (hàng f(x) của bảng biến thiên
  // để trống cho mũi tên chạy liền mạch như bảng vẽ tay).
  function buildRow(label, nodeAt, gapAt, opts = {}) {
    let n = 0;
    let g = 0;
    const cells = kinds.map((kind, i) => {
      const isNode = kind === 'node';
      const width = isNode ? nodeW : gapW;
      const content = isNode ? nodeAt(n++) : gapAt(g++);
      const borders = opts.openSides
        ? { left: i === 0 ? 'single' : 'nil', right: i === kinds.length - 1 ? 'single' : 'nil' }
        : null;
      return tc(width, content ? textPara(content, { jc: 'center', spaceAfter: 0 }) : '', {
        borders,
      });
    });
    const head = tc(labelW, textPara(label, { jc: 'center', italic: true, spaceAfter: 0 }));
    return tr(head + cells.join(''), opts.height);
  }

  const blank = () => '';
  let rows = buildRow(cfg.varName || 'x', (i) => nodes[i], blank);

  if (cfg.mode === 'bbt') {
    rows += buildRow(cfg.derivName || "f'(x)", (i) => marks[i], (i) => signs[i]);
    rows += buildRow(cfg.fnName || 'f(x)', blank, (i) => ARROW[signs[i]] || '', {
      height: 900,
      openSides: true,
    });
  } else {
    rows += buildRow(cfg.fnName || 'f(x)', (i) => marks[i], (i) => signs[i]);
  }

  return wrapBody(tbl(widths, rows, true));
}
