// Symbol/equation library for the Math Symbols task pane.
// type "unicode": inserted as plain text at the cursor.
// type "omml": inserted as a real Word equation object (OOXML/OMML).

const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function wrapOMath(inner) {
  return `<w:p ${W_NS} ${MATH_NS}><m:oMathPara><m:oMath>${inner}</m:oMath></m:oMathPara></w:p>`;
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
        name: 'Phân số a/b',
        preview: 'a⁄b',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>`
        ),
      },
      {
        name: 'Căn bậc hai',
        preview: '√x',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:rad><m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>`
        ),
      },
      {
        name: 'Căn bậc n',
        preview: 'ⁿ√x',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:rad><m:radPr><m:ctrlPr/></m:radPr><m:deg><m:r><m:t>n</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>`
        ),
      },
      {
        name: 'Lũy thừa xⁿ',
        preview: 'xⁿ',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>n</m:t></m:r></m:sup></m:sSup>`
        ),
      },
      {
        name: 'Chỉ số dưới xᵢ',
        preview: 'xᵢ',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>`
        ),
      },
      {
        name: 'Tích phân có cận',
        preview: '∫ₐᵇ f(x)dx',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/><m:ctrlPr/></m:naryPr><m:sub><m:r><m:t>a</m:t></m:r></m:sub><m:sup><m:r><m:t>b</m:t></m:r></m:sup><m:e><m:r><m:t>f(x)dx</m:t></m:r></m:e></m:nary>`
        ),
      },
      {
        name: 'Tổng Σ có cận',
        preview: '∑ᵢ₌₁ⁿ aᵢ',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:ctrlPr/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>a_i</m:t></m:r></m:e></m:nary>`
        ),
      },
      {
        name: 'Giới hạn',
        preview: 'lim x→a f(x)',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:func><m:funcPr><m:ctrlPr/></m:funcPr><m:fName><m:limLow><m:limLowPr><m:ctrlPr/></m:limLowPr><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>x→a</m:t></m:r></m:lim></m:limLow></m:fName><m:e><m:r><m:t>f(x)</m:t></m:r></m:e></m:func>`
        ),
      },
      {
        name: 'Ma trận 2x2',
        preview: '[a b; c d]',
        type: 'omml',
        ooxml: wrapOMath(
          `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/><m:ctrlPr/></m:dPr><m:e><m:m><m:mPr><m:baseJc m:val="center"/><m:ctrlPr/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m></m:e></m:d>`
        ),
      },
    ],
  },
];
