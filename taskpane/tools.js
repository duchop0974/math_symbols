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

// Chỉ dẫn hiện khi trỏ chuột. Tra theo nhãn ô hoặc chữ trên nút nên không phải
// sửa từng chỗ gọi, và nhãn nào đổi mà quên chỉ dẫn thì chỉ mất tooltip.
const HINTS = {
  // --- định dạng
  'Phông': 'Phông cho mọi thứ add-in chèn ra. Đề thi ở Việt Nam dùng Times New Roman; công thức toán vẫn là Cambria Math theo quy định của Word.',
  'Cỡ chữ (pt)': 'Đề thi thường cỡ 12. Văn bản hành chính theo Nghị định 30/2020 là 13–14.',
  'Lề trang': 'Chỉ dùng để tính bề rộng bảng cho khớp cột chữ. Lề thật của tài liệu phải đặt trong Word: Layout → Margins → Custom Margins.',
  'Giãn dòng': 'Đề thi thường để Đơn cho gọn trang. Văn bản hành chính yêu cầu ít nhất 1,5 dòng.',

  // --- đầu đề
  'Cơ quan quản lý': 'Dòng trên cùng bên trái, vd SỞ GD&ĐT HÀ NỘI hoặc PHÒNG GD&ĐT BÌNH XUYÊN. Để trống thì không có dòng này.',
  'Trường / đơn vị': 'Dòng thứ hai bên trái, in đậm. Để trống thì bỏ qua.',
  'Nhãn đề': 'In đậm bên trái, vd ĐỀ CHÍNH THỨC hoặc ĐỀ THAM KHẢO.',
  'Số trang': 'Sinh ra dòng "(Đề thi có N trang)". Để trống thì không có dòng này.',
  'Kỳ thi': 'Dòng đầu cột phải, in đậm, vd KỲ THI TỐT NGHIỆP THPT NĂM 2025.',
  'Lớp': 'Ghép với Năm học thành dòng "LỚP 12, NĂM HỌC 2024-2025". Điền một trong hai cũng được.',
  'Năm học': 'Ghép với Lớp thành một dòng. Để trống cả hai thì không có dòng này.',
  'Nhãn dòng môn': 'Chữ đứng trước tên môn: "Môn thi" như đề Bộ, hay "MÔN" như đề học sinh giỏi.',
  'Môn': 'Tên môn, tự viết hoa khi chèn.',
  'Thời gian (phút)': 'Chỉ điền số, sinh ra dòng "Thời gian làm bài: 90 phút".',
  'Ghi chú thời gian': 'Đặt trong ngoặc ngay sau số phút, vd "không kể thời gian phát đề" hoặc "không kể thời gian giao đề".',
  'Mã đề': 'Sinh ra dòng "Mã đề thi 101" bên phải. Để trống thì đề không có mã đề.',
  'Dòng họ tên & số báo danh': 'Thêm hai dòng chỗ trống cho thí sinh ghi họ tên và số báo danh, ngay dưới đầu đề.',

  // --- tuỳ chọn khung đề
  'Kèm đầu đề thi': 'Bỏ tick nếu tài liệu đã có sẵn đầu đề và chỉ cần dựng phần câu hỏi.',
  'Kèm dòng kết đề': 'Dòng "…Hết…" kèm lời dặn thí sinh và chỗ ghi họ tên ở cuối đề.',
  'Đánh số liên tục cả đề': 'Bật: đánh 1, 2, 3… xuyên suốt cả đề. Tắt: mỗi phần đánh lại từ câu 1, đúng cấu trúc đề Bộ GD&ĐT.',
  'Có tiêu đề phần': 'Tắt nếu đề không chia PHẦN I, II, III — thường gặp ở đề tự luận và đề học sinh giỏi.',
  'Tiêu đề phần ngắn gọn': 'Bỏ cụm nêu tên dạng câu, chỉ còn "PHẦN I. Thí sinh trả lời từ câu 1 đến câu 12. …".',
  'Kèm bảng hướng dẫn chấm': 'Thêm bảng chấm vào cuối tài liệu, bắt đầu ở một trang mới.',
  'Số dòng bảng chấm': 'Số dòng để trống của bảng hướng dẫn chấm kèm theo.',
  '+ Thêm phần': 'Thêm một phần mới vào cuối danh sách.',
  'Khôi phục mặc định': 'Về lại 12 – 4 – 6 theo cấu trúc đề Bộ GD&ĐT từ 2025.',
  'Dựng vào Word': 'Chèn cả khung đề vào vị trí con trỏ, một lần duy nhất.',

  // --- tab Chèn
  'Loại bảng': 'Bảng biến thiên có 3 dòng (x, đạo hàm, hàm số); bảng xét dấu chỉ 2 dòng.',
  'Tên biến': 'Chữ ở đầu dòng thứ nhất, thường là x.',
  'Tên hàm': 'Chữ ở đầu dòng cuối, thường là f(x) hoặc y.',
  'Tên đạo hàm': 'Chữ ở đầu dòng giữa của bảng biến thiên, thường là f′(x).',
  'Nghiệm / điểm đặc biệt': 'Các mốc trên trục, cách nhau bằng dấu phẩy, vd: -1, 0, 2.',
  'Kèm −∞ và +∞': 'Thêm hai đầu mút vô cực vào hai bên bảng.',
  'Chèn bảng vào Word': 'Chèn bảng đúng như phần Xem trước ở trên.',
  'Đầu đề thi': 'Chèn riêng đầu đề vào vị trí con trỏ, dùng thông tin đã nhập ở tab Tạo đề.',
  'Dòng kết đề (…Hết…)': 'Chèn riêng khối kết đề vào vị trí con trỏ.',
  'Số dòng trống': 'Số dòng để trống của bảng sắp chèn.',
  'Câu | Ý | Nội dung | Điểm': 'Bảng hướng dẫn chấm 4 cột, có cột Ý cho các ý a, b, c.',
  'Câu | Nội dung | Điểm': 'Bảng 3 cột, thường dùng cho bảng cấu trúc đề.',

  // --- tab Hoàn thiện
  'Đáp án': 'Nhận "1A 2B 3C", "1.A, 2.B", "ABCD…" hoặc đáp án đúng/sai kiểu "1 ĐSSĐ".',
  'Số câu mỗi dòng': 'Bảng đáp án xuống dòng sau bao nhiêu câu.',
  'Câu bắt đầu': 'Số câu đầu tiên, dùng khi đáp án nhập vào chỉ có chữ cái mà không có số.',
  'Số câu (bảng trống)': 'Số cột của bảng đáp án trống sắp chèn.',
  'Chèn bảng đáp án': 'Chèn bảng đáp án từ chuỗi đã nhập ở ô Đáp án.',
  'Bảng trống': 'Chèn bảng đáp án chưa điền, để tự viết tay.',
  'Đánh số lại từ': 'Số bắt đầu khi đánh số lại toàn bộ câu trong tài liệu.',
  'Đánh số lại các câu': 'Quét cả tài liệu, đánh số lại mọi đoạn bắt đầu bằng "Câu n." theo thứ tự hiện tại. Chỉ chạy trong Word.',
  'Trộn đề gốc': 'Mã đề cách nhau bằng dấu phẩy, vd: 101, 102, 103, 104.',
  'Đáp án đề gốc': 'Nhập để add-in tính luôn bảng đáp án cho từng mã đề sau khi trộn.',
  'Trộn thành các mã đề': 'Giữ nguyên đề gốc, nối thêm từng mã đề vào cuối tài liệu, mỗi mã một trang mới. Lưu tài liệu trước khi bấm. Chỉ chạy trong Word.',
};

// Nhãn ô nằm ở <span> đầu tiên của .field; nút thì lấy đúng chữ trên nút.
function applyHints(root) {
  root.querySelectorAll('.field').forEach((field) => {
    const label = field.firstChild ? field.firstChild.textContent : '';
    const hint = HINTS[label];
    if (!hint) return;
    field.title = hint;
    field.querySelectorAll('input, select, textarea').forEach((el) => {
      el.title = hint;
    });
  });
  root.querySelectorAll('.act').forEach((btn) => {
    const hint = HINTS[btn.textContent];
    if (hint) btn.title = hint;
  });
}

// ---------------------------------------------------------------- tab Đề thi

const PARTS_KEY = 'mathSymbols.parts';

const KIND_OPTIONS = [
  ['mc', 'Trắc nghiệm A–D'],
  ['tf', 'Đúng / Sai'],
  ['sa', 'Trả lời ngắn'],
  ['tl', 'Tự luận'],
];

// Nhãn gọn cho dòng tóm tắt — không viết thường nhãn dropdown vì "A–D" sẽ thành "a–d".
const KIND_SHORT = {
  mc: 'trắc nghiệm',
  tf: 'đúng/sai',
  sa: 'trả lời ngắn',
  tl: 'tự luận',
};

// Mặc định theo cấu trúc đề Bộ GD&ĐT từ 2025 (12 – 4 – 6); đề khác thì thêm/bớt.
const defaultParts = () => [
  { kind: 'mc', count: 12, cols: 2 },
  { kind: 'tf', count: 4 },
  { kind: 'sa', count: 6 },
];

function loadParts() {
  try {
    const saved = JSON.parse(localStorage.getItem(PARTS_KEY));
    if (Array.isArray(saved) && saved.length) return saved;
  } catch {
    // hỏng localStorage thì quay về mặc định
  }
  return defaultParts();
}

function saveParts(parts) {
  try {
    localStorage.setItem(PARTS_KEY, JSON.stringify(parts));
  } catch {
    // không lưu được thì chỉ mất phần nhớ giữa các lần mở
  }
}

// Mỗi ô là một dòng của đầu đề; để trống thì dòng đó không được chèn.
const HEADER_FIELDS = [
  ['so', 'Cơ quan quản lý', 'SỞ GD&ĐT ... / PHÒNG GD&ĐT ...'],
  ['truong', 'Trường / đơn vị', 'TRƯỜNG THPT ...'],
  ['deLabel', 'Nhãn đề', 'ĐỀ CHÍNH THỨC'],
  ['soTrang', 'Số trang', '04'],
  ['kyThi', 'Kỳ thi', 'KỲ THI TỐT NGHIỆP THPT NĂM 2025'],
  ['khoi', 'Lớp', '12'],
  ['namHoc', 'Năm học', '2024-2025'],
  ['monLabel', 'Nhãn dòng môn', 'Môn thi'],
  ['mon', 'Môn', 'Toán'],
  ['thoiGian', 'Thời gian (phút)', '90'],
  ['ghiChuGio', 'Ghi chú thời gian', 'không kể thời gian phát đề'],
  ['maDe', 'Mã đề', '101'],
];

// Mồi sẵn cho lần mở đầu tiên, để bấm dựng là ra đầu đề dùng được ngay; những ô
// riêng của từng trường (cơ quan, kỳ thi) để trống cho giáo viên tự điền.
const DEFAULT_HEADER = {
  deLabel: 'ĐỀ CHÍNH THỨC',
  monLabel: 'Môn thi',
  mon: 'Toán',
  thoiGian: '90',
  ghiChuGio: 'không kể thời gian phát đề',
  hoTen: true,
};

function loadHeader() {
  try {
    return JSON.parse(localStorage.getItem(HEADER_KEY)) || DEFAULT_HEADER;
  } catch {
    return DEFAULT_HEADER;
  }
}

function readHeaderForm() {
  const data = { hoTen: !!(document.getElementById('hdr-hoTen') || {}).checked };
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

function renderBuildPanel(panel) {
  const saved = loadHeader();
  let parts = loadParts();

  const check = (id, checked) => {
    const box = h('input', { id, type: 'checkbox' });
    box.checked = checked;
    return box;
  };

  // ---------------------------------------------------------------- định dạng

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

  const formatRows = [
    fieldRow('Phông', fontSel),
    fieldRow('Cỡ chữ (pt)', sizeBox),
    fieldRow('Lề trang', marginSel),
    fieldRow('Giãn dòng', lineSel),
    h('p', { id: 'page-width', class: 'measure' }),
    note(
      'Lề trang ở đây chỉ để tính bề rộng bảng cho khớp cột chữ — add-in không đổi được lề ' +
        'của tài liệu, thầy/cô đặt trong Word: Layout → Margins → Custom Margins. Phông áp cho ' +
        'mọi thứ add-in chèn ra ở cả ba tab; công thức toán vẫn dùng Cambria Math.'
    ),
  ];

  // ------------------------------------------------------------- đầu đề thi

  // Mỗi ô là một dòng của đầu đề; để trống thì dòng đó không xuất hiện, nên cùng
  // một biểu mẫu ra được đề Bộ, đề Sở, đề trường hay đề học sinh giỏi.
  const headerRows = HEADER_FIELDS.map(([key, label, placeholder]) =>
    fieldRow(label, input(`hdr-${key}`, saved[key], { placeholder }))
  );
  const hoTenBox = check('hdr-hoTen', saved.hoTen !== false);
  headerRows.push(fieldRow('Dòng họ tên & số báo danh', hoTenBox));

  // -------------------------------------------------------- các phần của đề

  const partsBox = h('div', { id: 'skel-parts', class: 'parts' });
  const briefOn = () => !!(document.getElementById('skel-brief') || {}).checked;
  const titlesOn = () => (document.getElementById('skel-titles') || {}).checked !== false;

  function partExtras(part) {
    if (part.kind === 'tl') {
      const diem = input('', part.diem || '2,0', { placeholder: '2,0' });
      diem.className = 'tiny';
      diem.title = 'Điểm ghi trong ngoặc, vd "Câu 1 (2,0 điểm).".';
      diem.addEventListener('input', () => {
        part.diem = diem.value;
        touchParts();
      });
      const subs = input('', String(part.subs || 0), { type: 'number', min: 0, max: 5 });
      subs.className = 'tiny';
      subs.title = 'Số ý a), b), c) của mỗi câu. Để 0 thì câu không chia ý.';
      subs.addEventListener('input', () => {
        part.subs = parseInt(subs.value, 10) || 0;
        touchParts();
      });
      return [
        h('span', { class: 'tiny-label', text: 'điểm' }),
        diem,
        h('span', { class: 'tiny-label', text: 'ý' }),
        subs,
      ];
    }
    if (part.kind !== 'mc') return [];
    const cols = select(
      [
        ['2', '2 cột'],
        ['1', '1 cột'],
        ['4', '4 cột'],
      ],
      String(part.cols || 2),
      (e) => {
        part.cols = parseInt(e.target.value, 10);
        touchParts();
      }
    );
    cols.className = 'tiny';
    cols.title = 'Xếp 4 phương án A, B, C, D thành mấy cột.';
    return [cols];
  }

  function renderParts() {
    partsBox.innerHTML = '';
    parts.forEach((part, i) => {
      const kindSel = select(KIND_OPTIONS, part.kind, (e) => {
        part.kind = e.target.value;
        renderParts();
        touchParts();
      });
      kindSel.title = 'Dạng câu hỏi của phần này.';
      const count = input('', String(part.count), { type: 'number', min: 0, max: 99 });
      count.className = 'tiny';
      count.title = 'Số câu của phần này. Để 0 thì bỏ hẳn phần, kể cả dòng tiêu đề.';
      count.addEventListener('input', () => {
        part.count = parseInt(count.value, 10) || 0;
        touchParts();
      });

      const row = h('div', { class: 'part-row' }, [
        h('span', { class: 'part-no', text: ROMAN[i] || String(i + 1) }),
        kindSel,
        count,
        h('span', { class: 'tiny-label', text: 'câu' }),
      ]
        .concat(partExtras(part))
        .concat([
          h('button', {
            class: 'icon',
            text: '⤓',
            title: 'Chỉ chèn riêng phần này',
            onclick: () =>
              Pane.insertOoxml(
                partBlock(part, i, 1, briefOn(), titlesOn()),
                `Đã chèn phần ${ROMAN[i] || i + 1}`
              ),
          }),
          h('button', {
            class: 'icon',
            text: '✕',
            title: 'Bỏ phần này',
            onclick: () => {
              parts.splice(i, 1);
              renderParts();
              touchParts();
            },
          }),
        ]));

      const title = h('input', {
        type: 'text',
        class: 'part-title',
        placeholder: 'Tiêu đề tự sinh theo chuẩn — gõ để thay bằng chữ của mình',
        title:
          'Để trống thì tự sinh theo câu chữ đề Bộ GD&ĐT từ 2025, tự cập nhật khi đổi số câu. ' +
          'Gõ vào để thay bằng chữ của mình.',
      });
      title.value = part.title || '';
      title.addEventListener('input', () => {
        part.title = title.value.trim();
        touchParts();
      });

      partsBox.appendChild(h('div', { class: 'part-item' }, [row, title]));
    });
  }

  function touchParts() {
    saveParts(parts);
    skelSummary();
  }

  function skelConfig() {
    return {
      header: (document.getElementById('skel-header') || {}).checked !== false,
      footer: !!(document.getElementById('skel-footer') || {}).checked,
      continuous: !!(document.getElementById('skel-continuous') || {}).checked,
      brief: briefOn(),
      titles: titlesOn(),
      grading: (document.getElementById('skel-grading') || {}).checked
        ? { rows: Math.max(1, num('skel-grading-rows', 10)), sub: true }
        : null,
      parts,
    };
  }

  function skelSummary() {
    const box = document.getElementById('skel-summary');
    if (!box) return;
    const cfg = skelConfig();
    const bits = [];
    if (cfg.header) bits.push('đầu đề');
    parts.forEach((p, i) => {
      if (!p.count) return;
      bits.push(`Phần ${ROMAN[i] || i + 1} ${p.count} câu ${KIND_SHORT[p.kind] || p.kind}`);
    });
    if (cfg.footer) bits.push('dòng kết đề');
    if (cfg.grading) bits.push(`bảng hướng dẫn chấm ${cfg.grading.rows} dòng`);
    const total = parts.reduce((a, p) => a + (p.count || 0), 0);
    box.textContent = bits.length
      ? `Sẽ dựng: ${bits.join(' + ')} — tổng ${total} câu, đánh số ${
          cfg.continuous ? 'liên tục cả đề' : 'lại từ 1 mỗi phần'
        }.`
      : 'Chưa chọn gì để dựng.';
  }

  // Mục chính, đứng đầu: mọi thứ để ra một đề hoàn chỉnh nằm gọn trong đây.
  // Cả tab này là việc dựng khung đề, nên không bọc thêm một lớp gấp/mở nữa:
  // chỉ mục Định dạng — đặt một lần rồi thôi — được gấp lại.
  panel.appendChild(section('Định dạng', formatRows, false));

  panel.appendChild(h('div', { class: 'sub-head', text: 'Đầu đề thi' }));
  headerRows.forEach((row) => panel.appendChild(row));

  panel.appendChild(h('div', { class: 'sub-head', text: 'Các phần của đề' }));
  panel.appendChild(partsBox);
  panel.appendChild(
    h('div', { class: 'btn-row' }, [
      button('+ Thêm phần', () => {
        parts.push({ kind: 'mc', count: 10, cols: 2 });
        renderParts();
        touchParts();
      }),
      button('Khôi phục mặc định', () => {
        parts = defaultParts();
        renderParts();
        touchParts();
      }),
    ])
  );

  panel.appendChild(h('div', { class: 'sub-head', text: 'Tuỳ chọn' }));
  [
    fieldRow('Kèm đầu đề thi', check('skel-header', true)),
    fieldRow('Kèm dòng kết đề', check('skel-footer', true)),
    fieldRow('Đánh số liên tục cả đề', check('skel-continuous', false)),
    fieldRow('Có tiêu đề phần', check('skel-titles', true)),
    fieldRow('Tiêu đề phần ngắn gọn', check('skel-brief', false)),
    fieldRow('Kèm bảng hướng dẫn chấm', check('skel-grading', false)),
    fieldRow('Số dòng bảng chấm', input('skel-grading-rows', '10', { type: 'number', min: 1, max: 60 })),
  ].forEach((row) => panel.appendChild(row));

  panel.appendChild(h('p', { id: 'skel-summary', class: 'measure' }));
  panel.appendChild(
    button(
      'Dựng vào Word',
      () => {
        const cfg = skelConfig();
        const total = parts.reduce((a, p) => a + (p.count || 0), 0);
        if (!total && !cfg.header && !cfg.footer) {
          Pane.toast('Chưa chọn phần nào để dựng.', true);
          return;
        }
        Pane.insertOoxml(examSkeleton(readHeaderForm(), cfg), 'Đã dựng cả khung đề');
      },
      true
    )
  );
  panel.appendChild(
    note(
      'Ô đầu đề nào để trống thì dòng đó không xuất hiện — bỏ trống Mã đề là đề không có mã. ' +
        'Thêm bao nhiêu phần cũng được, mỗi phần tự chọn dạng câu và số câu; nút ⤓ chèn riêng ' +
        'một phần. Chèn từng câu thì dùng thanh nút nhanh ở trên cùng.'
    )
  );
  panel.addEventListener('change', skelSummary);

  renderParts();
  savePageSetting();
  skelSummary();
  applyHints(panel);
}

// ------------------------------------------------------------------ tab Chèn

// Mọi thứ chèn thêm vào khi đang gõ nội dung nằm ở đây: ký hiệu, bảng biến
// thiên, và các khối rời của đề. Đây là chỗ duy nhất chèn một khối đơn lẻ.
function renderInsertPanel(panel) {
  const symbols = h('div');
  panel.appendChild(section('Ký hiệu & công thức', [symbols], true));
  Pane.renderSymbols(symbols);

  const bbtBox = h('div');
  panel.appendChild(section('Bảng biến thiên / xét dấu', [bbtBox], false));
  renderVariationPanel(bbtBox);

  panel.appendChild(
    section('Khối rời của đề', [
      h('div', { class: 'btn-row' }, [
        button('Đầu đề thi', () =>
          Pane.insertOoxml(examHeader(readHeaderForm()), 'Đã chèn đầu đề thi')
        ),
        button('Dòng kết đề (…Hết…)', () =>
          Pane.insertOoxml(examFooter(), 'Đã chèn phần kết đề')
        ),
      ]),
      h('div', { class: 'sub-head', text: 'Bảng hướng dẫn chấm' }),
      fieldRow('Số dòng trống', input('hdc-rows', '10', { type: 'number', min: 1, max: 60 })),
      h('div', { class: 'btn-row' }, [
        button('Câu | Ý | Nội dung | Điểm', () =>
          Pane.insertOoxml(
            gradingTable(Math.max(1, num('hdc-rows', 10)), true),
            'Đã chèn bảng hướng dẫn chấm'
          )
        ),
        button('Câu | Nội dung | Điểm', () =>
          Pane.insertOoxml(
            gradingTable(Math.max(1, num('hdc-rows', 10)), false),
            'Đã chèn bảng cấu trúc đề'
          )
        ),
      ]),
      note(
        'Dùng khi đề đã dựng xong mà cần thêm một khối. Muốn có sẵn ngay từ đầu thì bật ô ' +
          'tương ứng ở tab Tạo đề. Chèn từng câu thì dùng thanh nút nhanh ở trên cùng.'
      ),
    ])
  );

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

  applyHints(panel);
}

// -------------------------------------------------- bảng biến thiên

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
  applyHints(panel);
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
  applyHints(panel);
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

// Ba tab theo ba giai đoạn làm việc: dựng khung → gõ nội dung → hoàn thiện.
const TOOL_TABS = [
  { name: 'Tạo đề', render: renderBuildPanel },
  { name: 'Chèn', render: renderInsertPanel },
  { name: 'Hoàn thiện', render: renderAnswerPanel },
];
