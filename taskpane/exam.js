// exam.js — dựng OOXML cho các khối phục vụ soạn đề thi:
// đầu đề thi, câu hỏi theo cấu trúc đề THPT, bảng biến thiên/xét dấu, bảng đáp án.
// Tất cả đều đi qua wrapBody() của symbols.js để thành gói Flat OPC hoàn chỉnh.

// Bề ngang vùng soạn thảo, tính bằng twip (1/20 pt) — bảng chèn ra phải vừa đúng
// cột chữ, nên phải khớp lề trang thật của tài liệu.
// Mặc định theo Nghị định 30/2020/NĐ-CP: A4 210mm, lề trái 30mm, lề phải 20mm.
const A4_WIDTH_MM = 210;
const MM_TO_TWIP = 1440 / 25.4;

const PAGE = { width: Math.round((A4_WIDTH_MM - 30 - 20) * MM_TO_TWIP) };

function setPageMargins(leftMm, rightMm) {
  const usable = A4_WIDTH_MM - leftMm - rightMm;
  if (usable > 40) PAGE.width = Math.round(usable * MM_TO_TWIP);
  return PAGE.width;
}

// Giãn dòng: 240 twip = đơn. NĐ 30/2020 yêu cầu ít nhất 1,5 dòng cho văn bản
// hành chính, còn đề thi thường để đơn cho gọn trang.
const LINE = { value: 240 };

function setLineSpacing(multiple) {
  LINE.value = Math.round(240 * multiple);
}

function esc(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Phông dùng cho mọi thứ add-in chèn ra. Đề thi ở Việt Nam mặc định Times New
// Roman 12; không đặt gì thì Word lấy phông Normal của tài liệu (thường Calibri).
const FONT = { name: 'Times New Roman', size: 12 };

function setFont(name, size) {
  if (name) FONT.name = name;
  if (size) FONT.size = size;
}

// Thứ tự con của w:rPr theo schema: rFonts → b → i → sz → szCs.
// w:sz tính theo nửa point, nên 12pt = 24.
function fontProps(opts = {}) {
  const half = Math.round((opts.size || FONT.size) * 2);
  const family = esc(FONT.name);
  return (
    `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>` +
    (opts.bold ? '<w:b/>' : '') +
    (opts.italic ? '<w:i/>' : '') +
    `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`
  );
}

function run(content, opts = {}) {
  return (
    `<w:r><w:rPr>${fontProps(opts)}</w:rPr>` +
    `<w:t xml:space="preserve">${esc(content)}</w:t></w:r>`
  );
}

// Thứ tự các con của w:pPr phải đúng theo schema OOXML (spacing → ind → jc),
// sai thứ tự thì Word từ chối cả gói với thông báo "problem with its contents".
// w:rPr trong w:pPr là định dạng của dấu kết đoạn — phải đặt phông ở đây nữa,
// nếu không giáo viên gõ tiếp vào đoạn trống sẽ ra phông mặc định của Word.
function spacingXml(opts) {
  const attrs = [];
  if (opts.spaceAfter !== undefined) attrs.push(`w:after="${opts.spaceAfter}"`);
  if (LINE.value !== 240) attrs.push(`w:line="${LINE.value}" w:lineRule="auto"`);
  return attrs.length ? `<w:spacing ${attrs.join(' ')}/>` : '';
}

function para(runsXml, opts = {}) {
  // Dấu kết đoạn chỉ mang phông và cỡ chữ, cố tình bỏ đậm/nghiêng: nếu nó đậm
  // thì giáo viên gõ tiếp ở cuối đoạn sẽ ra chữ đậm ngoài ý muốn.
  const ppr =
    spacingXml(opts) +
    (opts.ind ? `<w:ind w:left="${opts.ind}"/>` : '') +
    (opts.jc ? `<w:jc w:val="${opts.jc}"/>` : '') +
    `<w:rPr>${fontProps({ size: opts.size })}</w:rPr>`;
  return `<w:p><w:pPr>${ppr}</w:pPr>${runsXml}</w:p>`;
}

const textPara = (content, opts = {}) => para(run(content, opts), opts);

// "Câu 1." in đậm rồi một dấu cách thường: gõ tiếp sau đó ra chữ không đậm.
const stemPara = (label, opts = {}) =>
  para(run(label, { ...opts, bold: true }) + run(' ', { size: opts.size }), opts);

// Đoạn trống vẫn phải mang phông, nên không dùng <w:p/> trần.
const emptyPara = (opts) => para('', opts);

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
    `${contentXml || emptyPara()}</w:tc>`
  );
}

function tr(cellsXml, height) {
  const props = height ? `<w:trPr><w:trHeight w:val="${height}"/></w:trPr>` : '';
  return `<w:tr>${props}${cellsXml}</w:tr>`;
}

// Ngắt trang — dùng cho bảng hướng dẫn chấm và cho từng mã đề khi trộn đề.
// Là hàm chứ không phải hằng, vì phông có thể đổi sau khi script đã nạp; đoạn
// ngắt trang cũng phải mang phông như mọi đoạn khác.
const pageBreak = () => para(`<w:r><w:rPr>${fontProps()}</w:rPr><w:br w:type="page"/></w:r>`);

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
    `${rowsXml}</w:tbl>${emptyPara()}`
  );
}

// ---------------------------------------------------------------- đầu đề thi

// Đầu đề ghép từ các dòng rời, ô nào để trống thì dòng đó biến mất — nhờ vậy một
// biểu mẫu duy nhất ra được cả đề Bộ, đề Sở, đề trường lẫn đề học sinh giỏi,
// không phải chọn "kiểu đề" nào cả.
function examHeaderBody(f) {
  const half = Math.floor(PAGE.width / 2);
  const centred = { jc: 'center', spaceAfter: 0 };
  const line = (text, opts) => (text ? textPara(text, { ...centred, ...opts }) : '');

  const left =
    line(f.so) +
    line(f.truong, { bold: true }) +
    line(f.deLabel, { bold: true }) +
    line(f.soTrang ? `(Đề thi có ${f.soTrang} trang)` : '', { italic: true });

  const lop = [f.khoi ? `LỚP ${f.khoi}` : '', f.namHoc ? `NĂM HỌC ${f.namHoc}` : '']
    .filter(Boolean)
    .join(', ');
  const gio = f.thoiGian
    ? `Thời gian làm bài: ${f.thoiGian} phút` + (f.ghiChuGio ? ` (${f.ghiChuGio})` : '')
    : '';
  const right =
    line(f.kyThi, { bold: true }) +
    line(lop, { bold: true }) +
    line(f.mon ? `${f.monLabel || 'Môn thi'}: ${f.mon.toUpperCase()}` : '', { bold: true }) +
    line(gio, { italic: true });

  let out =
    left || right
      ? tbl(
          [half, half],
          tr(tc(half, left, { vAlign: 'top' }) + tc(half, right, { vAlign: 'top' })),
          false
        )
      : '';

  if (f.maDe) out += textPara(`Mã đề thi ${f.maDe}`, { jc: 'right', bold: true });
  if (f.hoTen) {
    out +=
      textPara('Họ, tên thí sinh: ................................................................', {
        spaceAfter: 0,
      }) + textPara('Số báo danh: ................................................................');
  }
  return out + emptyPara();
}

// Khối kết đề tự luận: dòng Hết, lời dặn và chỗ ghi họ tên thí sinh.
function examFooterBody() {
  return (
    textPara('…………..Hết…………', { jc: 'center', bold: true }) +
      textPara(
        'Thí sinh không được sử dụng máy tính cầm tay. Cán bộ coi thi không giải thích gì thêm.',
        { jc: 'center', italic: true }
      ) +
      textPara(
        'Họ và tên thí sinh: ..........................................................; ' +
          'Số báo danh: ....................'
      )
  );
}

// ---------------------------------------------------------------- câu hỏi

// Câu chữ lấy đúng theo đề tham khảo/chính thức của Bộ GD&ĐT từ 2025, nhưng số
// phần và dạng câu của mỗi phần là do người dùng đặt chứ không cố định.
const KIND_TEXT = {
  mc: {
    name: 'Câu trắc nghiệm nhiều phương án lựa chọn.',
    tail: ' Mỗi câu hỏi thí sinh chỉ chọn một phương án.',
  },
  tf: {
    name: 'Câu trắc nghiệm đúng sai.',
    tail: ' Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.',
  },
  sa: { name: 'Câu trắc nghiệm trả lời ngắn.', tail: '' },
  tl: { name: 'Tự luận.', tail: '' },
};

const ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'];

// index tính từ 0; from/to là số câu thật sự của phần, phụ thuộc cách đánh số.
function partTitle(index, kind, from, to, brief) {
  const k = KIND_TEXT[kind] || KIND_TEXT.mc;
  const kindText = brief ? '' : `${k.name} `;
  const range =
    from === to
      ? `Thí sinh trả lời câu ${from}.`
      : `Thí sinh trả lời từ câu ${from} đến câu ${to}.`;
  return `PHẦN ${ROMAN[index] || index + 1}. ${kindText}${range}${k.tail}`;
}

const MC_LABELS = ['A.', 'B.', 'C.', 'D.'];
const TF_LABELS = ['a)', 'b)', 'c)', 'd)'];

function multipleChoice(no, cols) {
  const stem = stemPara(`Câu ${no}.`, { spaceAfter: 0 });
  if (cols === 1) {
    const lines = MC_LABELS.map((l) => textPara(`${l} `, { ind: 284, spaceAfter: 0 }));
    return stem + lines.join('') + emptyPara();
  }
  const width = Math.floor(PAGE.width / cols);
  const cells = MC_LABELS.map((l) =>
    tc(width, textPara(`${l} `, { spaceAfter: 0 }), { vAlign: 'top' })
  );
  const rows = [];
  for (let i = 0; i < cells.length; i += cols) rows.push(tr(cells.slice(i, i + cols).join('')));
  return stem + tbl(new Array(cols).fill(width), rows.join(''), false);
}

function trueFalse(no) {
  return (
    stemPara(`Câu ${no}.`, { spaceAfter: 0 }) +
    TF_LABELS.map((l) => textPara(`${l} `, { ind: 284, spaceAfter: 0 })).join('') +
    emptyPara()
  );
}

function shortAnswer(no) {
  return stemPara(`Câu ${no}.`, { spaceAfter: 0 }) + emptyPara();
}

const SUB_LABELS = ['a)', 'b)', 'c)', 'd)', 'e)'];

// Câu tự luận kiểu đề HSG: "Câu 1 (4,0 điểm)." rồi các ý a), b), c) mỗi ý một dòng.
function essayQuestion(no, opts) {
  const diem = (opts && opts.diem) || '2,0';
  const subs = (opts && opts.subs) || 0;
  let out = stemPara(`Câu ${no} (${diem} điểm).`, { spaceAfter: 0 });
  for (let i = 0; i < subs; i += 1) {
    out += textPara(`${SUB_LABELS[i]} `, { ind: 284, spaceAfter: 0 });
  }
  return out + emptyPara();
}

const QUESTION_BUILDERS = { mc: multipleChoice, tf: trueFalse, sa: shortAnswer };

function questionBody(kind, start, count, cols) {
  const build = QUESTION_BUILDERS[kind];
  let out = '';
  for (let i = 0; i < count; i += 1) out += build(start + i, cols);
  return out;
}

function essayBody(start, count, opts) {
  let out = '';
  for (let i = 0; i < count; i += 1) out += essayQuestion(start + i, opts);
  return out;
}

// Bảng hướng dẫn chấm (Câu | Ý | Nội dung | Điểm) và bảng cấu trúc đề
// (Câu | Nội dung | Điểm) — hai bảng mọi đề HSG đều phải kèm.
function gradingBody(rowCount, withSubColumn) {
  const heads = withSubColumn ? ['Câu', 'Ý', 'Nội dung', 'Điểm'] : ['Câu', 'Nội dung', 'Điểm'];
  const narrow = 700;
  const points = 1100;
  const wide = PAGE.width - points - narrow * (withSubColumn ? 2 : 1);
  const widths = withSubColumn
    ? [narrow, narrow, wide, points]
    : [narrow, wide, points];

  const headRow = tr(
    heads
      .map((label, i) => tc(widths[i], textPara(label, { jc: 'center', bold: true })))
      .join('')
  );
  let rows = headRow;
  for (let i = 0; i < rowCount; i += 1) {
    rows += tr(widths.map((w) => tc(w, '')).join(''), 400);
  }
  return wrapBody(tbl(widths, rows, true));
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
  const cellW = Math.floor((PAGE.width - labelW) / perRow);
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

// Bảng đáp án tổng hợp: mỗi dòng một mã đề, mỗi cột một câu.
// rows: [{ code, entries: [{ no, ans }] }]
function keyMatrixTable(rows) {
  const count = rows.reduce((m, r) => Math.max(m, r.entries.length), 0);
  const labelW = 900;
  const cellW = Math.max(320, Math.floor((PAGE.width - labelW) / Math.max(1, count)));
  const widths = [labelW, ...new Array(count).fill(cellW)];

  let out = tr(
    tc(labelW, textPara('Mã đề', { jc: 'center', bold: true })) +
      Array.from({ length: count }, (_, i) =>
        tc(cellW, textPara(String(i + 1), { jc: 'center', bold: true }))
      ).join('')
  );
  rows.forEach((r) => {
    out += tr(
      tc(labelW, textPara(String(r.code), { jc: 'center', bold: true })) +
        Array.from({ length: count }, (_, i) =>
          tc(cellW, textPara((r.entries[i] || {}).ans || '', { jc: 'center' }))
        ).join('')
    );
  });

  return wrapBody(
    textPara('ĐÁP ÁN CÁC MÃ ĐỀ', { jc: 'center', bold: true }) + tbl(widths, out, true)
  );
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
    Math.floor((PAGE.width - labelW - nodeW * nodes.length) / signs.length)
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

// ------------------------------------------------- các khối chèn được (gói OPC)

const gradingTable = (rowCount, withSub) => wrapBody(gradingBody(rowCount, withSub));

const examHeader = (f) => wrapBody(examHeaderBody(f));
const examFooter = () => wrapBody(examFooterBody());
const questionBlock = (kind, start, count, cols) => wrapBody(questionBody(kind, start, count, cols));
const essayBlock = (start, count, opts) => wrapBody(essayBody(start, count, opts));
// Một phần của đề: dạng câu, số câu, tiêu đề (rỗng = không chèn tiêu đề).
// part = { kind, count, cols, diem, subs, title }
function partBody(part, index, from, brief, withTitle) {
  const to = from + part.count - 1;
  // Ô tiêu đề để trống nghĩa là dùng câu chữ tự sinh theo chuẩn; muốn bỏ hẳn
  // dòng tiêu đề thì tắt ở mức cả đề (withTitle = false).
  const title = part.title || partTitle(index, part.kind, from, to, brief);
  const head = withTitle === false ? '' : textPara(title, { bold: true });
  const body =
    part.kind === 'tl'
      ? essayBody(from, part.count, { diem: part.diem, subs: part.subs })
      : questionBody(part.kind, from, part.count, part.cols || 2);
  return head + body;
}

const partBlock = (part, index, from, brief, withTitle) =>
  wrapBody(partBody(part, index, from, brief, withTitle));

// Dựng cả bộ xương của đề trong MỘT gói: đầu đề, các phần theo đúng danh sách
// người dùng đặt, rồi dòng kết. Chèn một lần nhanh hơn hẳn bấm lần lượt.
// cfg = { header, footer, continuous, brief, parts: [...] }
function examSkeleton(f, cfg) {
  let out = cfg.header === false ? '' : examHeaderBody(f);
  let no = 1;

  cfg.parts.forEach((part, i) => {
    if (!part.count) return;
    out += partBody(part, i, no, cfg.brief, cfg.titles);
    // Đề Bộ đánh số lại từ 1 ở mỗi phần; đề tự luận thường đánh liên tục cả đề.
    no = cfg.continuous ? no + part.count : 1;
  });

  if (cfg.footer) out += examFooterBody();
  // Hướng dẫn chấm luôn sang trang mới để tách hẳn khỏi đề.
  if (cfg.grading) {
    out += pageBreak() + gradingBody(cfg.grading.rows, cfg.grading.sub);
  }
  return wrapBody(out);
}
