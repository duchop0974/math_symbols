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

// Lề trang theo Nghị định 30/2020/NĐ-CP: A4, trên/dưới 20–25mm, trái 30–35mm,
// phải 15–20mm. Bảng chèn ra phải khớp cột chữ nên phải biết lề thật của tài liệu.
const PAGE_KEY = 'mathSymbols.page';
const MARGIN_PRESETS = [
  ['30-20', 'Trái 30 – phải 20 mm (chuẩn NĐ 30)'],
  ['35-20', 'Trái 35 – phải 20 mm'],
  ['30-15', 'Trái 30 – phải 15 mm'],
  ['20-20', 'Lề 2 cm đều'],
];
const LINE_PRESETS = [
  ['1', 'Đơn'],
  ['1.15', '1,15'],
  ['1.5', '1,5 (chuẩn văn bản hành chính)'],
];

function applyPage(marginKey, lineValue) {
  const [left, right] = marginKey.split('-').map(Number);
  setPageMargins(left, right);
  setLineSpacing(parseFloat(lineValue));
}

function loadPageSetting() {
  try {
    const saved = JSON.parse(localStorage.getItem(PAGE_KEY));
    if (saved) applyPage(saved.margin, saved.line);
  } catch {
    // hỏng localStorage thì dùng mặc định chuẩn NĐ 30
  }
}

function savePageSetting() {
  const margin = val('page-margin') || '30-20';
  const line = val('page-line') || '1';
  applyPage(margin, line);
  try {
    localStorage.setItem(PAGE_KEY, JSON.stringify({ margin, line }));
  } catch {
    // không lưu được thì vẫn áp dụng cho phiên hiện tại
  }
  const out = document.getElementById('page-width');
  if (out) out.textContent = `Cột chữ rộng ${Math.round(PAGE.width / MM_TO_TWIP)} mm`;
}

// Khung gấp/mở được — task pane hẹp, mở hết mọi mục cùng lúc thì phải cuộn rất dài.
const OPEN_KEY = 'mathSymbols.openSections';

function openState() {
  try {
    return JSON.parse(localStorage.getItem(OPEN_KEY)) || {};
  } catch {
    return {};
  }
}

function section(title, children, defaultOpen) {
  const state = openState();
  const box = h('details', { class: 'tool-section' }, [h('summary', { text: title })].concat(children));
  if (state[title] === undefined ? defaultOpen : state[title]) box.setAttribute('open', '');
  box.addEventListener('toggle', () => {
    const next = openState();
    next[title] = box.open;
    try {
      localStorage.setItem(OPEN_KEY, JSON.stringify(next));
    } catch {
      // mất phần nhớ trạng thái gấp/mở thôi, không ảnh hưởng gì khác
    }
  });
  return box;
}

loadFontSetting();
loadPageSetting();

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

function steps(items) {
  return h(
    'ol',
    { class: 'steps' },
    items.map((s) => h('li', { text: s }))
  );
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

  const fontSel = select(
    FONT_CHOICES.map((f) => [f, f]),
    FONT.name,
    saveFontSetting
  );
  fontSel.id = 'font-name';
  const sizeBox = input('font-size', String(FONT.size), { type: 'number', min: 8, max: 20 });
  sizeBox.addEventListener('input', saveFontSetting);
  const marginSel = select(MARGIN_PRESETS, '30-20', savePageSetting);
  marginSel.id = 'page-margin';
  const lineSel = select(LINE_PRESETS, '1', savePageSetting);
  lineSel.id = 'page-line';
  try {
    const savedPage = JSON.parse(localStorage.getItem(PAGE_KEY));
    if (savedPage) {
      marginSel.value = savedPage.margin;
      lineSel.value = savedPage.line;
    }
  } catch {
    // giữ mặc định
  }

  panel.appendChild(
    section(
      'Định dạng chuẩn',
      [
        fieldRow('Phông', fontSel),
        fieldRow('Cỡ chữ (pt)', sizeBox),
        fieldRow('Lề trang', marginSel),
        fieldRow('Giãn dòng', lineSel),
        h('p', { id: 'page-width', class: 'measure' }),
        note(
          'Áp cho mọi thứ add-in chèn ra ở cả ba tab. Lề trang ở đây chỉ để tính bề rộng ' +
            'bảng cho khớp cột chữ — add-in không đổi được lề của tài liệu, thầy/cô đặt trong ' +
            'Word: Layout → Margins → Custom Margins. Công thức toán vẫn dùng Cambria Math.'
        ),
      ],
      false
    )
  );

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
  const rows = {};
  const headerRows = HEADER_FIELDS.map(([key, label, placeholder, only]) => {
    const row = fieldRow(label, input(`hdr-${key}`, saved[key], { placeholder }));
    rows[key] = { row, only };
    return row;
  });

  // Ẩn ô không thuộc kiểu đề đang chọn, và bật/tắt các khối chỉ hợp với một kiểu.
  function applyStyle() {
    Object.values(rows).forEach(({ row, only }) => {
      row.classList.toggle('hidden', !!only && only !== style);
    });
    panel.querySelectorAll('[data-only]').forEach((el) => {
      el.classList.toggle('hidden', el.dataset.only !== style);
    });
    skelSummary();
  }

  panel.appendChild(
    section(
      'Đầu đề thi',
      [fieldRow('Kiểu đề', styleSel)]
        .concat(headerRows)
        .concat([
          button('Chèn đầu đề thi', () =>
            Pane.insertOoxml(examHeader(readHeaderForm(), style), 'Đã chèn đầu đề thi')
          ),
        ]),
      true
    )
  );

  // Đề Bộ đánh số lại từ câu 1 ở mỗi phần, nên chèn tiêu đề phần thì đưa ô
  // "Bắt đầu từ câu" về 1 luôn cho khỏi phải sửa tay.
  const insertPart = (which) => {
    const count = Math.max(1, num(`part-${which}`, SECTION_COUNTS[which]));
    Pane.insertOoxml(
      sectionHeading(which, count, val('part-wording') === 'brief'),
      `Đã chèn tiêu đề Phần ${'I'.repeat(which)}`
    );
    const startBox = document.getElementById('q-start');
    if (startBox) startBox.value = '1';
  };

  const wordingSel = select(
    [
      ['full', 'Đầy đủ (có nêu tên dạng câu)'],
      ['brief', 'Ngắn gọn'],
    ],
    'full'
  );
  wordingSel.id = 'part-wording';

  const skelCols = select(
    [
      ['2', '2 cột'],
      ['1', '1 cột'],
      ['4', '4 cột'],
    ],
    '2'
  );
  skelCols.id = 'skel-cols';

  const check = (id, checked) => {
    const box = h('input', { id, type: 'checkbox' });
    box.checked = checked;
    return box;
  };

  // Mọi thứ quyết định khung đề nằm gọn ở đây, và tóm tắt cập nhật theo từng ô
  // để thấy trước sẽ dựng ra cái gì rồi mới bấm.
  const skelRows = {
    header: fieldRow('Kèm đầu đề thi', check('skel-header', true)),
    wording: fieldRow('Câu chữ tiêu đề phần', wordingSel),
    p1: fieldRow('Phần I — trắc nghiệm', input('part-1', '12', { type: 'number', min: 0, max: 99 })),
    p2: fieldRow('Phần II — đúng/sai', input('part-2', '4', { type: 'number', min: 0, max: 99 })),
    p3: fieldRow('Phần III — trả lời ngắn', input('part-3', '6', { type: 'number', min: 0, max: 99 })),
    cols: fieldRow('Xếp phương án', skelCols),
    essay: fieldRow('Số câu tự luận', input('part-essay', '9', { type: 'number', min: 0, max: 60 })),
    diem: fieldRow('Điểm mỗi câu', input('skel-diem', '2,0', { placeholder: '2,0' })),
    subs: fieldRow('Số ý mỗi câu', input('skel-subs', '0', { type: 'number', min: 0, max: 5 })),
    footer: fieldRow('Kèm dòng kết đề', check('skel-footer', true)),
  };
  ['p1', 'p2', 'p3', 'cols', 'wording'].forEach((k) => {
    skelRows[k].dataset.only = 'tn';
  });
  ['essay', 'diem', 'subs'].forEach((k) => {
    skelRows[k].dataset.only = 'tl';
  });

  function skelConfig() {
    const header = (document.getElementById('skel-header') || {}).checked !== false;
    const footer = !!(document.getElementById('skel-footer') || {}).checked;
    if (style === 'tl') {
      return {
        header,
        footer,
        essayCount: Math.max(0, num('part-essay', 9)),
        diem: val('skel-diem').trim() || '2,0',
        subs: num('skel-subs', 0),
      };
    }
    return {
      header,
      footer,
      counts: { 1: num('part-1', 12), 2: num('part-2', 4), 3: num('part-3', 6) },
      cols: num('skel-cols', 2),
      brief: val('part-wording') === 'brief',
    };
  }

  function skelSummary() {
    const box = document.getElementById('skel-summary');
    if (!box) return;
    const cfg = skelConfig();
    const bits = [];
    if (cfg.header) bits.push('đầu đề');
    if (style === 'tl') {
      if (cfg.essayCount) bits.push(`${cfg.essayCount} câu tự luận (${cfg.diem} điểm${cfg.subs ? `, ${cfg.subs} ý` : ''})`);
    } else {
      const parts = [1, 2, 3].filter((w) => cfg.counts[w]);
      parts.forEach((w) => bits.push(`Phần ${'I'.repeat(w)} ${cfg.counts[w]} câu`));
    }
    if (cfg.footer) bits.push('dòng kết đề');
    const total =
      style === 'tl' ? cfg.essayCount : [1, 2, 3].reduce((a, w) => a + cfg.counts[w], 0);
    box.textContent = bits.length
      ? `Sẽ dựng: ${bits.join(' + ')} — tổng ${total} câu.`
      : 'Chưa chọn gì để dựng.';
  }

  const skelSection = section(
    'Dựng cả khung đề',
    Object.values(skelRows).concat([
      h('p', { id: 'skel-summary', class: 'measure' }),
      button(
        'Dựng vào Word',
        () => {
          const cfg = skelConfig();
          const total =
            style === 'tl' ? cfg.essayCount : [1, 2, 3].reduce((a, w) => a + cfg.counts[w], 0);
          if (!total && !cfg.header && !cfg.footer) {
            Pane.toast('Chưa chọn phần nào để dựng.', true);
            return;
          }
          Pane.insertOoxml(examSkeleton(readHeaderForm(), style, cfg), 'Đã dựng cả khung đề');
        },
        true
      ),
      h('div', { class: 'sub-head', text: 'Hoặc chèn riêng tiêu đề từng phần', 'data-only': 'tn' }),
      h('div', { class: 'btn-row', 'data-only': 'tn' }, [
        button('PHẦN I', () => insertPart(1)),
        button('PHẦN II', () => insertPart(2)),
        button('PHẦN III', () => insertPart(3)),
      ]),
      note(
        'Câu chữ tiêu đề phần lấy theo đề tham khảo của Bộ GD&ĐT từ 2025 (Toán: 12 – 4 – 6 câu). ' +
          'Mỗi phần đánh số lại từ câu 1. Để 0 câu thì phần đó bị bỏ qua hẳn.'
      ),
    ]),
    true
  );
  skelSection.addEventListener('input', skelSummary);
  skelSection.addEventListener('change', skelSummary);
  panel.appendChild(skelSection);

  const startRow = fieldRow('Bắt đầu từ câu', input('q-start', '1', { type: 'number', min: 1 }));
  const countRow = fieldRow('Số câu chèn', input('q-count', '5', { type: 'number', min: 1, max: 100 }));
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
    Pane.insertOoxml(questionBlock(kind, start, count, num('skel-cols', 2)), `Đã chèn ${count} ${label}`);
    advance(start, count);
  };

  const insertEssay = () => {
    const { start, count } = range();
    const opts = { diem: val('skel-diem').trim() || '2,0', subs: num('skel-subs', 0) };
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
  ]);

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

  panel.appendChild(
    section(
      'Chèn câu hỏi',
      [startRow, countRow, essayButtons, mcButtons],
      true
    )
  );

  const hdc = section('Bảng hướng dẫn chấm', [
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
  hdc.dataset.only = 'tl';
  panel.appendChild(hdc);

  panel.appendChild(
    section('Nếu đề phải nộp bằng MathType', [
      steps([
        'Soạn xong cả đề bằng add-in này (công thức là Equation của Word).',
        'Word → tab MathType → Convert Equations.',
        'Nguồn: Word 2007 and later (OMML) equations — Phạm vi: Whole document — ' +
          'Đích: MathType equations (OLE objects).',
      ]),
      note(
        'Chuyển ở bước cuối cùng và giữ lại một bản chưa chuyển: sau khi thành MathType, ' +
          'muốn sửa công thức là phải mở MathType chứ không dùng lại add-in được. Bảng biến ' +
          'thiên, bảng đáp án và khung câu hỏi không bị ảnh hưởng.'
      ),
    ])
  );

  savePageSetting();
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

  panel.appendChild(heading('Đánh số câu'));
  panel.appendChild(fieldRow('Đánh số lại từ', input('renum-start', '1', { type: 'number', min: 1 })));
  panel.appendChild(
    button('Đánh số lại các câu', async () => {
      if (!wordAvailable()) return;
      try {
        const n = await renumberQuestions(Math.max(1, num('renum-start', 1)));
        Pane.toast(n ? `Đã đánh số lại ${n} câu.` : 'Không thấy đoạn nào bắt đầu bằng "Câu n.".', !n);
      } catch (e) {
        Pane.toast(`Lỗi: ${e.message}`, true);
      }
    })
  );

  panel.appendChild(heading('Trộn đề nhiều mã'));
  panel.appendChild(fieldRow('Mã đề', input('mix-codes', '101, 102, 103, 104', { placeholder: '101, 102' })));
  panel.appendChild(
    fieldRow(
      'Đáp án đề gốc',
      h('textarea', { id: 'mix-key', rows: '2', placeholder: '1A 2B 3C ... hoặc ABCD...' })
    )
  );
  panel.appendChild(
    note(
      'Đọc đề đang mở, giữ nguyên đề gốc và nối thêm từng mã đề vào cuối tài liệu, mỗi mã ' +
        'một trang mới. Trong mỗi phần, thứ tự câu bị đảo và thứ tự A/B/C/D trong từng câu ' +
        'cũng đảo; câu vẫn nằm đúng phần của nó. Nhập đáp án đề gốc thì được luôn bảng đáp án ' +
        'các mã. Hãy lưu tài liệu trước khi trộn.'
    )
  );
  panel.appendChild(
    button(
      'Trộn thành các mã đề',
      async () => {
        if (!wordAvailable()) return;
        const codes = val('mix-codes')
          .split(/[,;\s]+/)
          .map((c) => c.trim())
          .filter(Boolean);
        if (!codes.length) {
          Pane.toast('Chưa nhập mã đề nào.', true);
          return;
        }
        try {
          const plans = await makeVariants({ codes });
          lastPlans = plans;
          renderPlans(plans);
          const entries = parseAnswerKey(val('mix-key'), 1);
          if (entries.length) await appendKeyMatrix(plans, entries);
          Pane.toast(`Đã tạo ${plans.length} mã đề.`);
        } catch (e) {
          Pane.toast(`Lỗi trộn đề: ${e.message}`, true);
        }
      },
      true
    )
  );
  panel.appendChild(h('div', { id: 'shuffle-map' }));
}

let lastPlans = null;

// Đáp án chỉ tính cho phần đầu tiên có đủ 4 phương án — tức phần trắc nghiệm.
function keyRowsFor(plans, entries) {
  return plans.map(({ code, plan }) => {
    const part = plan.find((p) => p.perms.some(Boolean)) || plan[0];
    return { code, entries: remapKey(entries, part.order, part.perms) };
  });
}

async function appendKeyMatrix(plans, entries) {
  Pane.insertOoxml(keyMatrixTable(keyRowsFor(plans, entries)), 'Đã chèn bảng đáp án các mã đề');
}

function renderPlans(plans) {
  const box = document.getElementById('shuffle-map');
  if (!box) return;
  box.innerHTML = '';
  box.appendChild(h('div', { class: 'sub-head', text: 'Câu mới ← câu gốc (theo từng phần)' }));
  plans.forEach(({ code, plan }) => {
    const lines = plan
      .map((p, i) => `Phần ${i + 1}: ${p.order.map((o, n) => `${n + 1}←${o + 1}`).join(' ')}`)
      .join('\n');
    box.appendChild(h('div', { class: 'sub-head', text: `Mã đề ${code}` }));
    box.appendChild(h('pre', { class: 'out', text: lines }));
  });
}

// ---------------------------------------------------------------- đăng ký tab

const TOOL_TABS = [
  { name: 'Đề thi', render: renderExamPanel },
  { name: 'Bảng BT', render: renderVariationPanel },
  { name: 'Đáp án', render: renderAnswerPanel },
];
