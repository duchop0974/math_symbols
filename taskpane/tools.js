// tools.js — các tab công cụ soạn đề thi trong task pane.
// Mỗi tab tự dựng form riêng rồi gọi Pane.insertOoxml() (định nghĩa ở taskpane.js).

const HEADER_KEY = 'mathSymbols.examHeader';
const FONT_KEY = 'mathSymbols.font';

const FONT_CHOICES = ['Times New Roman', 'Cambria', 'Arial', 'Calibri'];

// Phông áp cho mọi khối add-in chèn ra, không riêng tab Đề thi — nên nạp ngay khi
// script chạy, kể cả khi giáo viên chưa mở tab đó lần nào.
function loadFontSetting() {
  try {
    const saved = JSON.parse(localStorage.getItem(FONT_KEY));
    if (saved) setFont(saved.name, saved.size);
  } catch {
    // hỏng localStorage thì dùng mặc định Times New Roman 12
  }
}

function saveFontSetting() {
  const name = val('font-name') || FONT.name;
  const size = parseFloat(val('font-size')) || FONT.size;
  setFont(name, size);
  try {
    localStorage.setItem(FONT_KEY, JSON.stringify({ name, size }));
  } catch {
    // không lưu được thì vẫn áp dụng cho phiên hiện tại
  }
}

loadFontSetting();

function h(tag, attrs, children) {
  const node = document.createElement(tag);
  Object.entries(attrs || {}).forEach(([k, v]) => {
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  });
  (children || []).forEach((c) => node.appendChild(c));
  return node;
}

function fieldRow(label, input) {
  return h('label', { class: 'field' }, [h('span', { text: label }), input]);
}

function input(id, value, opts) {
  const node = h('input', { id, type: (opts && opts.type) || 'text', value: value || '' });
  if (opts && opts.placeholder) node.placeholder = opts.placeholder;
  if (opts && opts.min !== undefined) node.min = opts.min;
  if (opts && opts.max !== undefined) node.max = opts.max;
  return node;
}

function button(label, onClick, primary) {
  return h('button', { class: primary ? 'act primary' : 'act', text: label, onclick: onClick });
}

function select(options, value, onChange) {
  const node = h('select', { onchange: onChange });
  options.forEach(([optValue, label]) => {
    node.appendChild(h('option', { value: optValue, text: label }));
  });
  // Phải đặt sau khi đã gắn đủ option: gán .selected cho option còn rời khỏi cây
  // DOM sẽ bị trình duyệt xoá khi option được chèn vào select.
  node.value = value;
  return node;
}

function heading(label) {
  return h('h4', { class: 'tool-head', text: label });
}

function note(label) {
  return h('p', { class: 'hint', text: label });
}

const val = (id) => (document.getElementById(id) || {}).value || '';
const num = (id, fallback) => {
  const n = parseInt(val(id), 10);
  return Number.isFinite(n) ? n : fallback;
};

// ---------------------------------------------------------------- tab Đề thi

const STYLE_KEY = 'mathSymbols.examStyle';

// Ô nào chỉ dùng cho một kiểu đề thì đánh dấu bằng `only`.
const HEADER_FIELDS = [
  ['so', 'Sở / Phòng GD&ĐT', 'PHÒNG GD&ĐT ...'],
  ['truong', 'Trường', 'TRƯỜNG THPT ...', 'tn'],
  ['kyThi', 'Kỳ thi', 'ĐỀ GIAO LƯU HỌC SINH GIỎI CẤP HUYỆN'],
  ['mon', 'Môn', 'Toán'],
  ['khoi', 'Lớp', '6'],
  ['namHoc', 'Năm học', '2022-2023', 'tl'],
  ['thoiGian', 'Thời gian (phút)', '120'],
  ['soTrang', 'Số trang', '01'],
  ['maDe', 'Mã đề', '101', 'tn'],
];

function loadStyle() {
  try {
    return localStorage.getItem(STYLE_KEY) || 'tl';
  } catch {
    return 'tl';
  }
}

function saveStyle(style) {
  try {
    localStorage.setItem(STYLE_KEY, style);
  } catch {
    // không lưu được thì chỉ mất phần nhớ giữa các lần mở
  }
}

function loadHeader() {
  try {
    return JSON.parse(localStorage.getItem(HEADER_KEY)) || {};
  } catch {
    return {};
  }
}

function readHeaderForm() {
  const data = {};
  HEADER_FIELDS.forEach(([key]) => {
    data[key] = val(`hdr-${key}`).trim();
  });
  try {
    localStorage.setItem(HEADER_KEY, JSON.stringify(data));
  } catch {
    // không lưu được thì vẫn chèn bình thường, chỉ mất phần nhớ giữa các lần mở
  }
  return data;
}

function renderExamPanel(panel) {
  const saved = loadHeader();

  panel.appendChild(heading('Phông chữ'));
  const fontSel = select(
    FONT_CHOICES.map((f) => [f, f]),
    FONT.name,
    saveFontSetting
  );
  fontSel.id = 'font-name';
  panel.appendChild(fieldRow('Phông', fontSel));
  const sizeBox = input('font-size', String(FONT.size), { type: 'number', min: 8, max: 20 });
  sizeBox.addEventListener('input', saveFontSetting);
  panel.appendChild(fieldRow('Cỡ chữ (pt)', sizeBox));
  panel.appendChild(
    note('Áp cho mọi thứ add-in chèn ra ở cả ba tab công cụ. Công thức toán vẫn dùng Cambria Math như Word quy định.')
  );

  panel.appendChild(heading('Đầu đề thi'));
  let style = loadStyle();
  const styleSel = select(
    [
      ['tl', 'Tự luận / học sinh giỏi'],
      ['tn', 'Trắc nghiệm (có mã đề)'],
    ],
    style,
    (e) => {
      style = e.target.value;
      saveStyle(style);
      applyStyle();
    }
  );
  panel.appendChild(fieldRow('Kiểu đề', styleSel));

  const rows = {};
  HEADER_FIELDS.forEach(([key, label, placeholder, only]) => {
    const row = fieldRow(label, input(`hdr-${key}`, saved[key], { placeholder }));
    rows[key] = { row, only };
    panel.appendChild(row);
  });

  // Ẩn ô không thuộc kiểu đề đang chọn, và bật/tắt các nút chỉ hợp với một kiểu.
  function applyStyle() {
    Object.values(rows).forEach(({ row, only }) => {
      row.classList.toggle('hidden', !!only && only !== style);
    });
    panel.querySelectorAll('[data-only]').forEach((el) => {
      el.classList.toggle('hidden', el.dataset.only !== style);
    });
  }

  panel.appendChild(
    button(
      'Chèn đầu đề thi',
      () => Pane.insertOoxml(examHeader(readHeaderForm(), style), 'Đã chèn đầu đề thi'),
      true
    )
  );

  const sections = h('div', { 'data-only': 'tn' }, [
    heading('Tiêu đề phần'),
    h('div', { class: 'btn-row' }, [
      button('PHẦN I', () => Pane.insertOoxml(sectionHeading(1), 'Đã chèn tiêu đề Phần I')),
      button('PHẦN II', () => Pane.insertOoxml(sectionHeading(2), 'Đã chèn tiêu đề Phần II')),
      button('PHẦN III', () => Pane.insertOoxml(sectionHeading(3), 'Đã chèn tiêu đề Phần III')),
    ]),
  ]);
  panel.appendChild(sections);

  panel.appendChild(heading('Chèn câu hỏi'));
  panel.appendChild(fieldRow('Bắt đầu từ câu', input('q-start', '1', { type: 'number', min: 1 })));
  panel.appendChild(fieldRow('Số câu chèn', input('q-count', '5', { type: 'number', min: 1, max: 100 })));
  const colsSel = select(
    [
      ['2', '2 cột'],
      ['1', '1 cột (mỗi dòng một phương án)'],
      ['4', '4 cột trên một dòng'],
    ],
    '2'
  );
  colsSel.id = 'q-cols';
  const colsRow = fieldRow('Xếp phương án', colsSel);
  colsRow.dataset.only = 'tn';
  panel.appendChild(colsRow);

  const diemRow = fieldRow('Điểm mỗi câu', input('q-diem', '2,0', { placeholder: '2,0' }));
  diemRow.dataset.only = 'tl';
  panel.appendChild(diemRow);
  const subsRow = fieldRow('Số ý (a, b, c...)', input('q-subs', '0', { type: 'number', min: 0, max: 5 }));
  subsRow.dataset.only = 'tl';
  panel.appendChild(subsRow);

  const advance = (start, count) => {
    const startBox = document.getElementById('q-start');
    if (startBox) startBox.value = String(start + count);
  };
  const range = () => ({
    start: Math.max(1, num('q-start', 1)),
    count: Math.min(100, Math.max(1, num('q-count', 1))),
  });

  const insertQuestions = (kind, label) => {
    const { start, count } = range();
    Pane.insertOoxml(questionBlock(kind, start, count, num('q-cols', 2)), `Đã chèn ${count} ${label}`);
    advance(start, count);
  };

  const insertEssay = () => {
    const { start, count } = range();
    const opts = { diem: val('q-diem').trim() || '2,0', subs: num('q-subs', 0) };
    Pane.insertOoxml(essayBlock(start, count, opts), `Đã chèn ${count} câu tự luận`);
    advance(start, count);
  };

  const essayButtons = h('div', { 'data-only': 'tl' }, [
    h('div', { class: 'btn-row' }, [
      button('Câu tự luận (có điểm)', insertEssay, true),
      button('Dòng kết đề (…Hết…)', () => Pane.insertOoxml(examFooter(), 'Đã chèn phần kết đề')),
    ]),
    note(
      'Chèn ra "Câu n (2,0 điểm)." rồi các ý a), b), c) mỗi ý một dòng — bấm vào sau dấu ' +
        'chấm để gõ nội dung. Ô "Bắt đầu từ câu" tự tăng sau mỗi lần chèn.'
    ),
    heading('Bảng hướng dẫn chấm'),
    fieldRow('Số dòng trống', input('hdc-rows', '10', { type: 'number', min: 1, max: 60 })),
    h('div', { class: 'btn-row' }, [
      button('Câu | Ý | Nội dung | Điểm', () =>
        Pane.insertOoxml(gradingTable(Math.max(1, num('hdc-rows', 10)), true), 'Đã chèn bảng hướng dẫn chấm')
      ),
      button('Câu | Nội dung | Điểm', () =>
        Pane.insertOoxml(gradingTable(Math.max(1, num('hdc-rows', 10)), false), 'Đã chèn bảng cấu trúc đề')
      ),
    ]),
  ]);
  panel.appendChild(essayButtons);

  const mcButtons = h('div', { 'data-only': 'tn' }, [
    h('div', { class: 'btn-row' }, [
      button('Trắc nghiệm A–D', () => insertQuestions('mc', 'câu trắc nghiệm'), true),
      button('Đúng / Sai', () => insertQuestions('tf', 'câu đúng/sai')),
      button('Trả lời ngắn', () => insertQuestions('sa', 'câu trả lời ngắn')),
    ]),
    note(
      'Câu hỏi chèn ra để trống phần nội dung — bấm vào sau "Câu n:" để gõ đề, ' +
        'sau mỗi nhãn A. B. C. D. để gõ phương án. Ô "Bắt đầu từ câu" tự tăng sau mỗi lần chèn.'
    ),
  ]);
  panel.appendChild(mcButtons);

  applyStyle();
}

// -------------------------------------------------- tab Bảng biến thiên

const bbt = { mode: 'bbt', signs: [], marks: [] };
const MARK_OPTIONS = [
  ['0', '0'],
  ['‖', '‖'],
  ['', 'trống'],
];

function bbtNodes() {
  const roots = val('bbt-roots')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  const infinite = document.getElementById('bbt-inf');
  if (!infinite || infinite.checked) return ['-∞', ...roots, '+∞'];
  return roots.length >= 2 ? roots : ['-∞', ...roots, '+∞'];
}

// Giữ lại lựa chọn cũ khi số nghiệm thay đổi, chỉ nới/cắt cho vừa.
function resize(list, length, fallback) {
  const out = list.slice(0, length);
  while (out.length < length) out.push(fallback);
  return out;
}

function refreshBbt() {
  const nodes = bbtNodes();
  bbt.signs = resize(bbt.signs, Math.max(1, nodes.length - 1), null).map((s, i) =>
    s || (i % 2 === 0 ? '+' : '-')
  );
  bbt.marks = resize(bbt.marks, nodes.length, '0');
  // Hai đầu mút ±∞ không có giá trị đạo hàm nên luôn để trống.
  nodes.forEach((n, i) => {
    if (n === '-∞' || n === '+∞') bbt.marks[i] = '';
  });

  const box = document.getElementById('bbt-dynamic');
  if (!box) return;
  box.innerHTML = '';

  box.appendChild(h('div', { class: 'sub-head', text: 'Dấu trên từng khoảng' }));
  const signRow = h('div', { class: 'chip-row' });
  bbt.signs.forEach((sign, i) => {
    const label = `${nodes[i]} → ${nodes[i + 1]}`;
    signRow.appendChild(
      h('label', { class: 'chip' }, [
        h('span', { text: label }),
        select(
          [
            ['+', '+'],
            ['-', '−'],
          ],
          sign,
          (e) => {
            bbt.signs[i] = e.target.value;
            refreshBbt();
          }
        ),
      ])
    );
  });
  box.appendChild(signRow);

  const inner = nodes
    .map((n, i) => ({ n, i }))
    .filter(({ n }) => n !== '-∞' && n !== '+∞');
  if (inner.length) {
    box.appendChild(h('div', { class: 'sub-head', text: 'Giá trị tại nghiệm' }));
    const markRow = h('div', { class: 'chip-row' });
    inner.forEach(({ n, i }) => {
      markRow.appendChild(
        h('label', { class: 'chip' }, [
          h('span', { text: `${bbt.mode === 'bbt' ? "f'" : 'f'}(${n})` }),
          select(MARK_OPTIONS, bbt.marks[i], (e) => {
            bbt.marks[i] = e.target.value;
            refreshBbt();
          }),
        ])
      );
    });
    box.appendChild(markRow);
  }

  renderBbtPreview(nodes);
}

function renderBbtPreview(nodes) {
  const box = document.getElementById('bbt-preview');
  if (!box) return;
  const cells = (nodeAt, gapAt) => {
    let out = '';
    nodes.forEach((_, i) => {
      out += `<td class="n">${nodeAt(i)}</td>`;
      if (i < bbt.signs.length) out += `<td class="g">${gapAt(i)}</td>`;
    });
    return out;
  };
  const blank = () => '';
  const arrow = (i) => (bbt.signs[i] === '+' ? '↗' : '↘');
  let rows = `<tr><th>${val('bbt-var') || 'x'}</th>${cells((i) => nodes[i], blank)}</tr>`;
  if (bbt.mode === 'bbt') {
    rows += `<tr><th>${val('bbt-deriv') || "f'(x)"}</th>${cells(
      (i) => bbt.marks[i] || '',
      (i) => (bbt.signs[i] === '-' ? '−' : '+')
    )}</tr>`;
    rows += `<tr class="tall"><th>${val('bbt-fn') || 'f(x)'}</th>${cells(blank, arrow)}</tr>`;
  } else {
    rows += `<tr><th>${val('bbt-fn') || 'f(x)'}</th>${cells(
      (i) => bbt.marks[i] || '',
      (i) => (bbt.signs[i] === '-' ? '−' : '+')
    )}</tr>`;
  }
  box.innerHTML = `<table class="bbt">${rows}</table>`;
}

function renderVariationPanel(panel) {
  panel.appendChild(heading('Bảng biến thiên / xét dấu'));

  const modeSel = select(
    [
      ['bbt', 'Bảng biến thiên (3 dòng)'],
      ['sign', 'Bảng xét dấu (2 dòng)'],
    ],
    bbt.mode,
    (e) => {
      bbt.mode = e.target.value;
      const derivField = document.getElementById('bbt-deriv-field');
      if (derivField) derivField.classList.toggle('hidden', bbt.mode !== 'bbt');
      refreshBbt();
    }
  );
  panel.appendChild(fieldRow('Loại bảng', modeSel));
  panel.appendChild(fieldRow('Tên biến', input('bbt-var', 'x')));
  panel.appendChild(fieldRow('Tên hàm', input('bbt-fn', 'f(x)')));

  const derivField = fieldRow('Tên đạo hàm', input('bbt-deriv', "f'(x)"));
  derivField.id = 'bbt-deriv-field';
  panel.appendChild(derivField);

  const roots = input('bbt-roots', '-1, 1', { placeholder: 'vd: -1, 1, 3' });
  roots.addEventListener('input', refreshBbt);
  panel.appendChild(fieldRow('Nghiệm / điểm đặc biệt', roots));

  const inf = h('input', { id: 'bbt-inf', type: 'checkbox', onchange: refreshBbt });
  inf.checked = true;
  panel.appendChild(fieldRow('Kèm −∞ và +∞', inf));

  ['bbt-var', 'bbt-fn', 'bbt-deriv'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', () => renderBbtPreview(bbtNodes()));
  });

  panel.appendChild(h('div', { id: 'bbt-dynamic' }));
  panel.appendChild(h('div', { class: 'sub-head', text: 'Xem trước' }));
  panel.appendChild(h('div', { id: 'bbt-preview', class: 'preview' }));

  panel.appendChild(
    button(
      'Chèn bảng vào Word',
      () => {
        const nodes = bbtNodes();
        Pane.insertOoxml(
          variationTable({
            mode: bbt.mode,
            varName: val('bbt-var') || 'x',
            fnName: val('bbt-fn') || 'f(x)',
            derivName: val('bbt-deriv') || "f'(x)",
            nodes,
            signs: bbt.signs.slice(0, Math.max(1, nodes.length - 1)),
            marks: bbt.marks.slice(0, nodes.length),
          }),
          'Đã chèn bảng'
        );
      },
      true
    )
  );
  panel.appendChild(
    note(
      'Bảng chèn ra là bảng Word thật: hàng f(x) để trống các ô để thầy/cô điền giá trị cực trị, ' +
        'mũi tên đã đặt sẵn theo dấu của đạo hàm. Kéo mép cột để chỉnh độ rộng nếu cần.'
    )
  );

  refreshBbt();
}

// ------------------------------------------------ tab Đáp án & trộn đề

let lastShuffleOrder = null; // order[viTriMoi] = viTriGoc

function shuffled(length) {
  const order = Array.from({ length }, (_, i) => i);
  for (let i = order.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [order[i], order[j]] = [order[j], order[i]];
  }
  return order;
}

function wordAvailable() {
  if (typeof Word === 'undefined' || !Pane.inWord) {
    Pane.toast('Chức năng này chỉ chạy khi task pane mở trong Word.', true);
    return false;
  }
  return true;
}

const QUESTION_RE = /^\s*Câu\s*(\d+)\s*([.:)])/;

async function renumberQuestions(start) {
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load('items/text');
    await ctx.sync();

    const jobs = [];
    paras.items.forEach((p) => {
      const m = QUESTION_RE.exec(p.text);
      if (m) jobs.push({ needle: m[0].trim(), sep: m[2], hits: p.search(m[0].trim(), { matchCase: true }) });
    });
    jobs.forEach((j) => j.hits.load('items'));
    await ctx.sync();

    jobs.forEach((j, i) => {
      if (j.hits.items.length) j.hits.items[0].insertText(`Câu ${start + i}${j.sep}`, 'Replace');
    });
    await ctx.sync();
    return jobs.length;
  });
}

async function shuffleQuestions() {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const paras = body.paragraphs;
    paras.load('items/text');
    await ctx.sync();

    const starts = [];
    paras.items.forEach((p, i) => {
      if (QUESTION_RE.test(p.text)) starts.push(i);
    });
    if (starts.length < 2) throw new Error('Cần ít nhất 2 đoạn bắt đầu bằng "Câu 1:" để trộn.');

    const last = paras.items.length - 1;
    const packages = starts.map((from, k) => {
      const to = k + 1 < starts.length ? starts[k + 1] - 1 : last;
      const range = paras.items[from].getRange('Whole').expandTo(paras.items[to].getRange('Whole'));
      return range.getOoxml();
    });
    // Toàn bộ vùng từ câu đầu tiên tới cuối tài liệu sẽ bị ghi đè bằng thứ tự mới.
    const region = paras.items[starts[0]].getRange('Whole').expandTo(paras.items[last].getRange('Whole'));
    await ctx.sync();

    const order = shuffled(packages.length);
    region.insertOoxml(packages[order[0]].value, 'Replace');
    await ctx.sync();
    for (let i = 1; i < order.length; i += 1) {
      body.insertOoxml(packages[order[i]].value, 'End');
      await ctx.sync(); // chèn tuần tự để giữ đúng thứ tự
    }
    return order;
  });
}

function renderAnswerPanel(panel) {
  panel.appendChild(heading('Bảng đáp án'));
  const key = h('textarea', { id: 'ans-key', rows: '3', placeholder: '1A 2B 3C 4D ... hoặc ABCDDCBA' });
  panel.appendChild(fieldRow('Đáp án', key));
  panel.appendChild(fieldRow('Số câu mỗi dòng', input('ans-per-row', '10', { type: 'number', min: 2, max: 20 })));
  panel.appendChild(fieldRow('Câu bắt đầu', input('ans-start', '1', { type: 'number', min: 1 })));
  panel.appendChild(fieldRow('Số câu (bảng trống)', input('ans-count', '40', { type: 'number', min: 1 })));
  panel.appendChild(
    h('div', { class: 'btn-row' }, [
      button(
        'Chèn bảng đáp án',
        () => {
          const entries = parseAnswerKey(val('ans-key'), num('ans-start', 1));
          if (!entries.length) {
            Pane.toast('Chưa nhập đáp án nào.', true);
            return;
          }
          Pane.insertOoxml(
            answerKeyTable(entries, num('ans-per-row', 10)),
            `Đã chèn bảng đáp án ${entries.length} câu`
          );
        },
        true
      ),
      button('Bảng trống', () => {
        const start = num('ans-start', 1);
        const count = Math.max(1, num('ans-count', 40));
        const entries = Array.from({ length: count }, (_, i) => ({ no: start + i, ans: '' }));
        Pane.insertOoxml(answerKeyTable(entries, num('ans-per-row', 10)), 'Đã chèn bảng đáp án trống');
      }),
    ])
  );

  panel.appendChild(heading('Đánh số & trộn đề'));
  panel.appendChild(fieldRow('Đánh số lại từ', input('renum-start', '1', { type: 'number', min: 1 })));
  panel.appendChild(
    button('Đánh số lại các câu', async () => {
      if (!wordAvailable()) return;
      try {
        const n = await renumberQuestions(Math.max(1, num('renum-start', 1)));
        Pane.toast(n ? `Đã đánh số lại ${n} câu.` : 'Không tìm thấy đoạn nào bắt đầu bằng "Câu n:".', !n);
      } catch (e) {
        Pane.toast(`Lỗi: ${e.message}`, true);
      }
    })
  );

  panel.appendChild(
    note(
      'Trộn đề đảo thứ tự các câu tính từ "Câu 1" đến hết tài liệu, rồi đánh số lại. ' +
        'Hãy lưu tài liệu trước khi trộn, và giữ bảng đáp án ở một file riêng.'
    )
  );
  panel.appendChild(
    button('Trộn thứ tự câu hỏi', async () => {
      if (!wordAvailable()) return;
      try {
        const order = await shuffleQuestions();
        lastShuffleOrder = order;
        await renumberQuestions(1);
        renderMapping(order);
        Pane.toast(`Đã trộn ${order.length} câu.`);
      } catch (e) {
        Pane.toast(`Lỗi trộn đề: ${e.message}`, true);
      }
    })
  );
  panel.appendChild(h('div', { id: 'shuffle-map' }));

  panel.appendChild(heading('Chuyển đáp án theo thứ tự mới'));
  panel.appendChild(
    fieldRow('Đáp án đề gốc', h('textarea', { id: 'remap-key', rows: '2', placeholder: 'ABCD... hoặc 1A 2B ...' }))
  );
  panel.appendChild(
    button('Chuyển đáp án', () => {
      if (!lastShuffleOrder) {
        Pane.toast('Chưa trộn đề trong phiên này nên chưa có thứ tự để chuyển.', true);
        return;
      }
      const entries = parseAnswerKey(val('remap-key'), 1);
      if (entries.length < lastShuffleOrder.length) {
        Pane.toast(`Đáp án gốc chỉ có ${entries.length} câu, cần ${lastShuffleOrder.length}.`, true);
        return;
      }
      const remapped = lastShuffleOrder.map((oldIdx, i) => ({ no: i + 1, ans: entries[oldIdx].ans }));
      const out = document.getElementById('remap-out');
      out.textContent = remapped.map((e) => `${e.no}${e.ans}`).join('  ');
      const box = document.getElementById('ans-key');
      if (box) box.value = out.textContent;
      Pane.toast('Đã chuyển đáp án và điền vào ô Bảng đáp án ở trên.');
    })
  );
  panel.appendChild(h('pre', { id: 'remap-out', class: 'out' }));
}

function renderMapping(order) {
  const box = document.getElementById('shuffle-map');
  if (!box) return;
  box.innerHTML = '';
  box.appendChild(h('div', { class: 'sub-head', text: 'Câu mới ← câu gốc' }));
  box.appendChild(
    h('pre', { class: 'out', text: order.map((oldIdx, i) => `${i + 1}←${oldIdx + 1}`).join('  ') })
  );
}

// ---------------------------------------------------------------- đăng ký tab

const TOOL_TABS = [
  { name: 'Đề thi', render: renderExamPanel },
  { name: 'Bảng BT', render: renderVariationPanel },
  { name: 'Đáp án', render: renderAnswerPanel },
];
