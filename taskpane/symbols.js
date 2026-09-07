// Symbol/equation library for the Math Symbols task pane.
// type "unicode": inserted as plain text at the cursor.
// type "omml": inserted as a real Word equation object (OOXML/OMML).

const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// Word vẽ ô trống điền được (placeholder) cho mọi slot OMML rỗng, nên các mẫu
// dưới đây cố tình để trống <m:num/>, <m:e/>... thay vì nhét sẵn chữ.
const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';

// Word mở dữ liệu Ooxml nhận được như một tài liệu tạm, nên nó phải là gói Flat
// OPC hoàn chỉnh — fragment <w:p> rời sẽ bị báo "problem with its contents".
// Bên trong dùng oMath (inline) chứ không phải oMathPara, để công thức nằm trong
// dòng văn bản và ghép được với công thức đang có.
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const DOC_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function wrapOMath(inner) {
  return (
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    '<pkg:part pkg:name="/_rels/.rels" ' +
    'pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
    `<pkg:xmlData><Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${DOC_REL}" Target="/word/document.xml"/>` +
    '</Relationships></pkg:xmlData></pkg:part>' +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_TYPE}">` +
    `<pkg:xmlData><w:document ${W_NS} ${MATH_NS}><w:body>` +
    `<w:p><m:oMath>${inner}</m:oMath></w:p>` +
    '</w:body></w:document></pkg:xmlData></pkg:part>' +
    '</pkg:package>'
  );
}

function tpl(name, preview, inner) {
  return { name, preview, type: 'omml', ooxml: wrapOMath(inner) };
}

// Chữ đứng (tên hàm, ký hiệu vi phân) — không in nghiêng như biến số.
function MTEXT(text) {
  return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${text}</m:t></m:r>`;
}

function DELIM(begChr, endChr, inner) {
  return (
    `<m:d><m:dPr><m:begChr m:val="${begChr}"/><m:endChr m:val="${endChr}"/>${CTRL}</m:dPr>` +
    `${inner}</m:d>`
  );
}

function NARY(chr, limLoc, hide = {}) {
  const sub = hide.subHide ? '<m:subHide m:val="1"/>' : '';
  const sup = hide.supHide ? '<m:supHide m:val="1"/>' : '';
  return (
    `<m:nary><m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="${limLoc}"/>${sub}${sup}${CTRL}` +
    '</m:naryPr><m:sub/><m:sup/><m:e/></m:nary>'
  );
}

function MATRIX(rows, cols) {
  const row = `<m:mr>${'<m:e/>'.repeat(cols)}</m:mr>`;
  return (
    `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/>` +
    `<m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs>${CTRL}</m:mPr>` +
    `${row.repeat(rows)}</m:m>`
  );
}

function EQARR(rows) {
  return `<m:eqArr><m:eqArrPr>${CTRL}</m:eqArrPr>${'<m:e/>'.repeat(rows)}</m:eqArr>`;
}

function FUNC(name) {
  return (
    `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName>${MTEXT(name)}</m:fName><m:e/></m:func>`
  );
}

function ACC(chr) {
  return `<m:acc><m:accPr><m:chr m:val="${chr}"/>${CTRL}</m:accPr><m:e/></m:acc>`;
}

function BAR(pos) {
  return `<m:bar><m:barPr><m:pos m:val="${pos}"/>${CTRL}</m:barPr><m:e/></m:bar>`;
}

function GROUPCHR(chr, pos, vertJc) {
  return (
    `<m:groupChr><m:groupChrPr><m:chr m:val="${chr}"/><m:pos m:val="${pos}"/>` +
    `<m:vertJc m:val="${vertJc}"/>${CTRL}</m:groupChrPr><m:e/></m:groupChr>`
  );
}

const SYMBOL_CATEGORIES = [
  {
    name: 'Hy Lạp',
    items: [
      ['α', 'alpha'], ['β', 'beta'], ['γ', 'gamma'], ['δ', 'delta'],
      ['ε', 'epsilon'], ['θ', 'theta'], ['λ', 'lambda'], ['μ', 'mu'],
      ['π', 'pi'], ['ρ', 'rho'], ['σ', 'sigma'], ['τ', 'tau'],
      ['φ', 'phi'], ['ω', 'omega'],
      ['Γ', 'Gamma'], ['Δ', 'Delta'], ['Θ', 'Theta'], ['Λ', 'Lambda'],
      ['Σ', 'Sigma'], ['Φ', 'Phi'], ['Ω', 'Omega'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Toán tử',
    items: [
      ['±', 'plus-minus'], ['∓', 'minus-plus'], ['×', 'times'], ['÷', 'divide'],
      ['≠', 'not equal'], ['≈', 'approx'], ['≡', 'equivalent'],
      ['≤', 'less or equal'], ['≥', 'greater or equal'], ['≪', 'much less'], ['≫', 'much greater'],
      ['√', 'square root'], ['∝', 'proportional to'], ['·', 'dot product'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Tập hợp',
    items: [
      ['∈', 'element of'], ['∉', 'not element of'], ['⊂', 'subset'], ['⊆', 'subset or equal'],
      ['⊃', 'superset'], ['⊇', 'superset or equal'], ['∪', 'union'], ['∩', 'intersection'],
      ['∅', 'empty set'], ['ℝ', 'real numbers'], ['ℕ', 'natural numbers'], ['ℤ', 'integers'],
      ['ℚ', 'rationals'], ['ℂ', 'complex numbers'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Logic',
    items: [
      ['∀', 'for all'], ['∃', 'there exists'], ['∄', 'does not exist'], ['¬', 'not'],
      ['∧', 'and'], ['∨', 'or'], ['⇒', 'implies'], ['⇔', 'iff'], ['⊕', 'xor'], ['⊢', 'entails'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Mũi tên',
    items: [
      ['→', 'right arrow'], ['←', 'left arrow'], ['↔', 'left-right arrow'],
      ['⇒', 'double right'], ['⇐', 'double left'], ['⇔', 'double left-right'],
      ['↑', 'up arrow'], ['↓', 'down arrow'], ['↦', 'maps to'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Giải tích',
    items: [
      ['∫', 'integral'], ['∬', 'double integral'], ['∮', 'contour integral'],
      ['∑', 'summation'], ['∏', 'product'], ['∂', 'partial derivative'],
      ['∇', 'nabla'], ['∞', 'infinity'], ['lim', 'limit'], ['°', 'degree'],
    ].map(([symbol, name]) => ({ symbol, name, type: 'unicode' })),
  },
  {
    name: 'Phân số & mũ',
    items: [
      tpl('Phân số', '▫⁄▫', `<m:f><m:fPr>${CTRL}</m:fPr><m:num/><m:den/></m:f>`),
      tpl(
        'Phân số nghiêng',
        '▫/▫',
        `<m:f><m:fPr><m:type m:val="skw"/>${CTRL}</m:fPr><m:num/><m:den/></m:f>`
      ),
      tpl(
        'Phân số một dòng',
        '▫÷▫',
        `<m:f><m:fPr><m:type m:val="lin"/>${CTRL}</m:fPr><m:num/><m:den/></m:f>`
      ),
      tpl(
        'Xếp chồng (không gạch)',
        '▫ ▫',
        `<m:f><m:fPr><m:type m:val="noBar"/>${CTRL}</m:fPr><m:num/><m:den/></m:f>`
      ),
      tpl(
        'Căn bậc hai',
        '√▫',
        `<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/><m:e/></m:rad>`
      ),
      tpl('Căn bậc n', 'ⁿ√▫', `<m:rad><m:radPr>${CTRL}</m:radPr><m:deg/><m:e/></m:rad>`),
      tpl('Lũy thừa', '▫^▫', `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e/><m:sup/></m:sSup>`),
      tpl('Chỉ số dưới', '▫_▫', `<m:sSub><m:sSubPr>${CTRL}</m:sSubPr><m:e/><m:sub/></m:sSub>`),
      tpl(
        'Chỉ số trên+dưới',
        '▫_▫^▫',
        `<m:sSubSup><m:sSubSupPr>${CTRL}</m:sSubSupPr><m:e/><m:sub/><m:sup/></m:sSubSup>`
      ),
      tpl(
        'Chỉ số đặt trước',
        '^▫_▫▫',
        `<m:sPre><m:sPrePr>${CTRL}</m:sPrePr><m:sub/><m:sup/><m:e/></m:sPre>`
      ),
      tpl(
        'Mũ của ngoặc',
        '(▫)^▫',
        `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${DELIM('(', ')', '<m:e/>')}</m:e><m:sup/></m:sSup>`
      ),
    ],
  },
  {
    name: 'Tổng & tích phân',
    items: [
      tpl('Tổng Σ có cận', '∑▫', NARY('∑', 'undOvr')),
      tpl('Tổng Σ một cận', '∑▫', NARY('∑', 'undOvr', { supHide: true })),
      tpl('Tích Π có cận', '∏▫', NARY('∏', 'undOvr')),
      tpl('Tích phân có cận', '∫▫', NARY('∫', 'subSup')),
      tpl('Tích phân không cận', '∫▫', NARY('∫', 'subSup', { subHide: true, supHide: true })),
      tpl('Tích phân hai lớp', '∬▫', NARY('∬', 'subSup', { subHide: true, supHide: true })),
      tpl('Tích phân đường', '∮▫', NARY('∮', 'subSup')),
      tpl('Hợp ⋃ có cận', '⋃▫', NARY('⋃', 'undOvr')),
      tpl('Giao ⋂ có cận', '⋂▫', NARY('⋂', 'undOvr')),
    ],
  },
  {
    name: 'Ngoặc & ma trận',
    items: [
      tpl('Ngoặc đơn', '(▫)', DELIM('(', ')', '<m:e/>')),
      tpl('Ngoặc vuông', '[▫]', DELIM('[', ']', '<m:e/>')),
      tpl('Ngoặc nhọn', '{▫}', DELIM('{', '}', '<m:e/>')),
      tpl('Ngoặc góc', '⟨▫⟩', DELIM('⟨', '⟩', '<m:e/>')),
      tpl('Trị tuyệt đối', '|▫|', DELIM('|', '|', '<m:e/>')),
      tpl('Chuẩn vector', '‖▫‖', DELIM('‖', '‖', '<m:e/>')),
      tpl(
        'Tổ hợp',
        '(▫ ▫)',
        DELIM(
          '(',
          ')',
          `<m:e><m:f><m:fPr><m:type m:val="noBar"/>${CTRL}</m:fPr><m:num/><m:den/></m:f></m:e>`
        )
      ),
      tpl('Hệ 2 phương trình', '{▫;▫', DELIM('{', '', `<m:e>${EQARR(2)}</m:e>`)),
      tpl('Hệ 3 phương trình', '{▫;▫;▫', DELIM('{', '', `<m:e>${EQARR(3)}</m:e>`)),
      tpl('Ma trận 2x2', '[▫ ▫]', DELIM('[', ']', `<m:e>${MATRIX(2, 2)}</m:e>`)),
      tpl('Ma trận 3x3', '[▫ ▫ ▫]', DELIM('[', ']', `<m:e>${MATRIX(3, 3)}</m:e>`)),
      tpl('Ma trận 2x2 ngoặc tròn', '(▫ ▫)', DELIM('(', ')', `<m:e>${MATRIX(2, 2)}</m:e>`)),
      tpl('Định thức 2x2', '|▫ ▫|', DELIM('|', '|', `<m:e>${MATRIX(2, 2)}</m:e>`)),
      tpl('Vector cột 2', '(▫;▫)', DELIM('(', ')', `<m:e>${MATRIX(2, 1)}</m:e>`)),
      tpl('Vector cột 3', '(▫;▫;▫)', DELIM('(', ')', `<m:e>${MATRIX(3, 1)}</m:e>`)),
    ],
  },
  {
    name: 'Hàm & dấu',
    items: [
      tpl(
        'Giới hạn',
        'lim ▫',
        `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName><m:limLow><m:limLowPr>${CTRL}</m:limLowPr><m:e>${MTEXT('lim')}</m:e><m:lim/></m:limLow></m:fName><m:e/></m:func>`
      ),
      tpl(
        'Max có điều kiện',
        'max ▫',
        `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName><m:limLow><m:limLowPr>${CTRL}</m:limLowPr><m:e>${MTEXT('max')}</m:e><m:lim/></m:limLow></m:fName><m:e/></m:func>`
      ),
      tpl('sin', 'sin ▫', FUNC('sin')),
      tpl('cos', 'cos ▫', FUNC('cos')),
      tpl('tan', 'tan ▫', FUNC('tan')),
      tpl('ln', 'ln ▫', FUNC('ln')),
      tpl(
        'log cơ số',
        'log_▫ ▫',
        `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName><m:sSub><m:sSubPr>${CTRL}</m:sSubPr><m:e>${MTEXT('log')}</m:e><m:sub/></m:sSub></m:fName><m:e/></m:func>`
      ),
      tpl(
        'Đạo hàm d▫/d▫',
        'd▫⁄d▫',
        `<m:f><m:fPr>${CTRL}</m:fPr><m:num>${MTEXT('d')}</m:num><m:den>${MTEXT('d')}</m:den></m:f>`
      ),
      tpl(
        'Đạo hàm riêng ∂▫/∂▫',
        '∂▫⁄∂▫',
        `<m:f><m:fPr>${CTRL}</m:fPr><m:num>${MTEXT('∂')}</m:num><m:den>${MTEXT('∂')}</m:den></m:f>`
      ),
      tpl('Vector (mũi tên trên)', '→ trên ▫', ACC('⃗')),
      tpl('Mũ (hat)', '^ trên ▫', ACC('̂')),
      tpl('Chấm trên', '· trên ▫', ACC('̇')),
      tpl('Gạch ngang trên', '‾▫', BAR('top')),
      tpl('Gạch ngang dưới', '_▫', BAR('bot')),
      tpl('Ngoặc nhọn trên', '⏞▫', GROUPCHR('⏞', 'top', 'bot')),
      tpl('Ngoặc nhọn dưới', '⏟▫', GROUPCHR('⏟', 'bot', 'top')),
    ],
  },
];
