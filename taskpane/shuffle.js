// shuffle.js — trộn đề thành nhiều mã đề.
//
// Phần lõi (tách gói, trộn phương án, tính lại đáp án) là hàm thuần trên chuỗi
// OOXML nên chạy và kiểm chứng được ngoài Word; chỉ phần đọc/ghi tài liệu mới
// cần Office.js.

const PART_RE = /^\s*PHẦN\s+[IVX]+/i;
const CAU_RE = /^\s*Câu\s*(\d+)\s*[.:)]/;
const OPTION_RE = /^\s*([A-D])[.)]\s*/;
const LETTERS = ['A', 'B', 'C', 'D'];

const xmlParser = new DOMParser();
const xmlSerializer = new XMLSerializer();

function parseXml(text) {
  const doc = xmlParser.parseFromString(text, 'application/xml');
  if (doc.querySelector('parsererror')) throw new Error('OOXML đọc được từ Word không hợp lệ.');
  return doc;
}

// ---------------------------------------------------------- ghép gói Flat OPC

// getOoxml() trả về nguyên một gói Flat OPC chứ không phải mảnh, nên muốn nối
// nhiều khối lại phải bóc lấy ruột <w:body> của từng gói.
const BODY_OPEN = /<w:body\b[^>]*>/;
const SECT_PR = /<w:sectPr[\s\S]*?<\/w:sectPr>\s*$|<w:sectPr\b[^>]*\/>\s*$/;

function splitPackage(pkgXml) {
  const open = BODY_OPEN.exec(pkgXml);
  const close = pkgXml.lastIndexOf('</w:body>');
  if (!open || close === -1) throw new Error('Không tìm thấy <w:body> trong gói OOXML.');
  const inner = pkgXml.slice(open.index + open[0].length, close);
  const sect = SECT_PR.exec(inner);
  return {
    head: pkgXml.slice(0, open.index + open[0].length),
    body: sect ? inner.slice(0, sect.index) : inner,
    sectPr: sect ? sect[0] : '',
    tail: pkgXml.slice(close),
  };
}

const bodyOf = (pkgXml) => splitPackage(pkgXml).body;

// Dùng lại chính gói Word trả về làm vỏ, chỉ thay ruột — nhờ vậy giữ nguyên
// styles.xml và các part khác, không mất định dạng theo style của tài liệu gốc.
function rebuildPackage(carrierPkgXml, bodyXml) {
  const p = splitPackage(carrierPkgXml);
  return p.head + bodyXml + p.sectPr + p.tail;
}

// ------------------------------------------------------------ trộn phương án

const unitText = (el) =>
  [...el.getElementsByTagName('w:t')].map((n) => n.textContent).join('');

const unitLetter = (el) => {
  const m = OPTION_RE.exec(unitText(el));
  return m ? m[1] : null;
};

// Phương án nằm trong ô bảng (kiểu 2/4 cột) hoặc mỗi phương án một đoạn.
function optionUnits(scope) {
  const cells = [...scope.getElementsByTagName('w:tc')].filter(unitLetter);
  if (cells.length === 4) return cells;
  const paras = [...scope.getElementsByTagName('w:p')].filter(unitLetter);
  return paras.length === 4 ? paras : [];
}

function setLabel(unit, letter) {
  const texts = [...unit.getElementsByTagName('w:t')];
  const target = texts.find((t) => OPTION_RE.test(t.textContent));
  if (!target) return false;
  target.textContent = target.textContent.replace(OPTION_RE, `${letter}. `);
  return true;
}

function randomPerm(n) {
  const p = Array.from({ length: n }, (_, i) => i);
  for (let i = p.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [p[i], p[j]] = [p[j], p[i]];
  }
  return p;
}

// perm[viTriMoi] = viTriCu. Trả về ruột body mới và perm đã dùng (null nếu câu
// này không có đủ 4 phương án — câu đúng/sai, trả lời ngắn, tự luận).
function shuffleOptions(bodyXml, perm) {
  const doc = parseXml(`<root ${W_NS} ${MATH_NS}>${bodyXml}</root>`);
  const units = optionUnits(doc);
  if (units.length !== 4) return { body: bodyXml, perm: null };

  const used = perm || randomPerm(4);
  const clones = units.map((u) => u.cloneNode(true));
  units.forEach((unit, newIdx) => {
    const src = clones[used[newIdx]];
    while (unit.firstChild) unit.removeChild(unit.firstChild);
    [...src.childNodes].forEach((n) => unit.appendChild(n.cloneNode(true)));
    setLabel(unit, LETTERS[newIdx]);
  });

  const root = doc.documentElement;
  const out = [...root.childNodes].map((n) => xmlSerializer.serializeToString(n)).join('');
  return { body: out, perm: used };
}

// ------------------------------------------------------------------- đáp án

// order[viTriMoi] = viTriCu của câu; perms[viTriCu] = hoán vị phương án của câu đó.
function remapKey(entries, order, perms) {
  return order.map((oldIdx, i) => {
    const src = entries[oldIdx];
    if (!src) return { no: i + 1, ans: '' };
    const perm = perms[oldIdx];
    if (!perm || src.ans.length !== 1) return { no: i + 1, ans: src.ans };
    const oldPos = LETTERS.indexOf(src.ans);
    if (oldPos === -1) return { no: i + 1, ans: src.ans };
    const newPos = perm.indexOf(oldPos);
    return { no: i + 1, ans: newPos === -1 ? src.ans : LETTERS[newPos] };
  });
}

// -------------------------------------------------- sửa chữ trong ruột body

// Đổi chữ ở <w:t> đầu tiên khớp `re`. Dùng để đánh lại số câu và thay mã đề mà
// không đụng tới phần còn lại (công thức, định dạng) của khối.
function replaceFirstText(bodyXml, re, make) {
  const doc = parseXml(`<root ${W_NS} ${MATH_NS}>${bodyXml}</root>`);
  const target = [...doc.getElementsByTagName('w:t')].find((t) => re.test(t.textContent));
  if (!target) return bodyXml;
  target.textContent = make(target.textContent);
  return [...doc.documentElement.childNodes]
    .map((n) => xmlSerializer.serializeToString(n))
    .join('');
}

const renumberBlock = (bodyXml, no) =>
  replaceFirstText(bodyXml, CAU_RE, (t) => t.replace(CAU_RE, (m, old) => m.replace(old, String(no))));

const setMaDe = (bodyXml, code) =>
  replaceFirstText(bodyXml, /Mã đề/, (t) => t.replace(/(Mã đề(?:\s*thi)?\s*:?\s*).*/, `$1${code}`));

// ----------------------------------------------------- dựng một mã đề từ gốc

// source: { headerBody, parts: [{ headingBody, questions: [body...] }] }
// Trả về ruột body của mã đề và kế hoạch trộn để tính lại đáp án.
function buildVariantBody(source, code) {
  let out = pageBreak() + (code ? setMaDe(source.headerBody, code) : source.headerBody);
  const plan = [];

  source.parts.forEach((part) => {
    const order = randomPerm(part.questions.length);
    const perms = new Array(part.questions.length).fill(null);
    let text = part.headingBody || '';
    order.forEach((oldIdx, newIdx) => {
      const res = shuffleOptions(part.questions[oldIdx]);
      perms[oldIdx] = res.perm;
      // Mỗi phần đánh số lại từ câu 1, đúng như cấu trúc đề Bộ.
      text += renumberBlock(res.body, newIdx + 1);
    });
    out += text;
    plan.push({ order, perms });
  });

  return { body: out, plan };
}

// ------------------------------------------------------- đọc / ghi tài liệu

// Đọc đề gốc: phần đầu đề, rồi từng phần với danh sách khối câu hỏi.
async function readSource(ctx) {
  const paras = ctx.document.body.paragraphs;
  paras.load('items/text');
  await ctx.sync();

  const items = paras.items;
  const marks = items.map((p) => (PART_RE.test(p.text) ? 'part' : CAU_RE.test(p.text) ? 'cau' : ''));
  const firstMark = marks.findIndex(Boolean);
  if (firstMark === -1) throw new Error('Không thấy đoạn nào bắt đầu bằng "Câu 1." để trộn.');

  const range = (a, b) =>
    items[a].getRange('Whole').expandTo(items[b].getRange('Whole')).getOoxml();

  const header = firstMark > 0 ? range(0, firstMark - 1) : null;
  const parts = [];
  let current = null;

  marks.forEach((mark, i) => {
    if (mark === 'part') {
      current = { heading: range(i, i), questions: [] };
      parts.push(current);
      return;
    }
    if (mark !== 'cau') return;
    if (!current) {
      current = { heading: null, questions: [] };
      parts.push(current);
    }
    let end = i;
    while (end + 1 < marks.length && !marks[end + 1]) end += 1;
    current.questions.push(range(i, end));
  });

  await ctx.sync();
  return { header, parts };
}

// Tạo cfg.codes mã đề, nối tiếp vào cuối tài liệu, mỗi mã một trang mới.
async function makeVariants(cfg) {
  return Word.run(async (ctx) => {
    const raw = await readSource(ctx);
    const carrier = (raw.header || raw.parts[0].heading || raw.parts[0].questions[0]).value;
    const source = {
      headerBody: raw.header ? bodyOf(raw.header.value) : '',
      parts: raw.parts.map((p) => ({
        headingBody: p.heading ? bodyOf(p.heading.value) : '',
        questions: p.questions.map((q) => bodyOf(q.value)),
      })),
    };
    if (!source.parts.some((p) => p.questions.length >= 2)) {
      throw new Error('Cần ít nhất 2 câu trong một phần thì mới trộn được.');
    }

    const plans = [];
    for (const code of cfg.codes) {
      const built = buildVariantBody(source, code);
      plans.push({ code, plan: built.plan });
      ctx.document.body.insertOoxml(rebuildPackage(carrier, built.body), 'End');
      await ctx.sync(); // chèn tuần tự để các mã đề giữ đúng thứ tự
    }
    return plans;
  });
}
