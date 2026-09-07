// Symbol/equation library for the Math Symbols task pane.
// type "unicode": inserted as plain text at the cursor.
// type "omml": inserted as a real Word equation object (OOXML/OMML).

const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// Word vẽ ô trống điền được (placeholder) cho mọi slot OMML rỗng, nên các mẫu
// dưới đây cố tình để trống <m:num/>, <m:e/>... thay vì nhét sẵn chữ.
const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';

// oMath (inline) chứ không phải oMathPara: công thức nằm trong dòng văn bản nên
// chèn tiếp mẫu khác khi con trỏ đang trong công thức sẽ ghép vào cùng vùng toán.
function wrapOMath(inner) {
  return `<w:p ${W_NS} ${MATH_NS}><m:oMath>${inner}</m:oMath></w:p>`;
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
    name: 'Mẫu công thức',
    items: [
      {
        name: 'Phân số',
        preview: '▫⁄▫',
        type: 'omml',
        ooxml: wrapOMath(`<m:f><m:fPr>${CTRL}</m:fPr><m:num/><m:den/></m:f>`),
      },
      {
        name: 'Căn bậc hai',
        preview: '√▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/><m:e/></m:rad>`
        ),
      },
      {
        name: 'Căn bậc n',
        preview: 'ⁿ√▫',
        type: 'omml',
        ooxml: wrapOMath(`<m:rad><m:radPr>${CTRL}</m:radPr><m:deg/><m:e/></m:rad>`),
      },
      {
        name: 'Lũy thừa',
        preview: '▫^▫',
        type: 'omml',
        ooxml: wrapOMath(`<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e/><m:sup/></m:sSup>`),
      },
      {
        name: 'Chỉ số dưới',
        preview: '▫_▫',
        type: 'omml',
        ooxml: wrapOMath(`<m:sSub><m:sSubPr>${CTRL}</m:sSubPr><m:e/><m:sub/></m:sSub>`),
      },
      {
        name: 'Chỉ số trên+dưới',
        preview: '▫_▫^▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:sSubSup><m:sSubSupPr>${CTRL}</m:sSubSupPr><m:e/><m:sub/><m:sup/></m:sSubSup>`
        ),
      },
      {
        name: 'Tích phân có cận',
        preview: '∫▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/>${CTRL}</m:naryPr><m:sub/><m:sup/><m:e/></m:nary>`
        ),
      },
      {
        name: 'Tổng Σ có cận',
        preview: '∑▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/>${CTRL}</m:naryPr><m:sub/><m:sup/><m:e/></m:nary>`
        ),
      },
      {
        name: 'Tích Π có cận',
        preview: '∏▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:nary><m:naryPr><m:chr m:val="∏"/><m:limLoc m:val="undOvr"/>${CTRL}</m:naryPr><m:sub/><m:sup/><m:e/></m:nary>`
        ),
      },
      {
        name: 'Giới hạn',
        preview: 'lim ▫',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName><m:limLow><m:limLowPr>${CTRL}</m:limLowPr><m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>lim</m:t></m:r></m:e><m:lim/></m:limLow></m:fName><m:e/></m:func>`
        ),
      },
      {
        name: 'Ngoặc đơn',
        preview: '(▫)',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/>${CTRL}</m:dPr><m:e/></m:d>`
        ),
      },
      {
        name: 'Ma trận 2x2',
        preview: '[▫ ▫; ▫ ▫]',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/>${CTRL}</m:dPr><m:e><m:m><m:mPr><m:baseJc m:val="center"/>${CTRL}</m:mPr><m:mr><m:e/><m:e/></m:mr><m:mr><m:e/><m:e/></m:mr></m:m></m:e></m:d>`
        ),
      },
    ],
  },
];
