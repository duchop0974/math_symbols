(function () {
  const FAV_KEY = 'mathSymbols.favorites';
  const RECENT_KEY = 'mathSymbols.recent';

  let activeCategory = TOOL_TABS[0].name;
  let activeGroup = SYMBOL_CATEGORIES[0].name;
  let favorites = loadFavorites();
  let recent = loadRecent();
  let inWord = false;

  function loadRecent() {
    try {
      return JSON.parse(localStorage.getItem(RECENT_KEY)) || [];
    } catch {
      return [];
    }
  }

  // Ghi lại mẫu vừa dùng để thanh nút nhanh luôn có sẵn thứ đang gõ dở.
  function pushRecent(item) {
    const key = keyOf(item);
    recent = [key].concat(recent.filter((k) => k !== key)).slice(0, 12);
    try {
      localStorage.setItem(RECENT_KEY, JSON.stringify(recent));
    } catch {
      // không lưu được thì danh sách chỉ sống trong phiên này
    }
  }

  function loadFavorites() {
    try {
      return JSON.parse(localStorage.getItem(FAV_KEY)) || [];
    } catch {
      return [];
    }
  }

  function saveFavorites() {
    try {
      localStorage.setItem(FAV_KEY, JSON.stringify(favorites));
    } catch {
      // storage unavailable — favorites stay in-memory for this session
    }
  }

  const keyOf = (item) => `${item.type}:${item.name}`;
  const isFav = (item) => favorites.includes(keyOf(item));

  function toggleFav(item) {
    const key = keyOf(item);
    const i = favorites.indexOf(key);
    if (i === -1) favorites.push(key);
    else favorites.splice(i, 1);
    saveFavorites();
    renderQuickSyms();
    if (refreshSymbolList) refreshSymbolList();
  }

  const allItems = () => SYMBOL_CATEGORIES.flatMap((c) => c.items);

  function send(data, isOmml, label) {
    if (!inWord) {
      navigator.clipboard?.writeText(data);
      toast(`Chế độ xem thử: đã copy "${label}" vào clipboard`);
      return;
    }

    Office.context.document.setSelectedDataAsync(
      data,
      { coercionType: isOmml ? Office.CoercionType.Ooxml : Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          toast(`Lỗi chèn: ${result.error.message}`, true);
        } else {
          toast(label);
        }
      }
    );
  }

  function insert(item) {
    const isOmml = item.type === 'omml';
    pushRecent(item);
    renderQuickSyms();
    send(
      isOmml ? item.ooxml : item.symbol,
      isOmml,
      inWord ? `Đã chèn ${item.symbol || item.name}` : item.name
    );
  }

  let toastTimer;
  function toast(message, isError) {
    const el = document.getElementById('toast');
    el.textContent = message;
    el.classList.toggle('error', !!isError);
    el.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.add('hidden'), 2000);
  }

  function buildButton(item) {
    const btn = document.createElement('button');
    btn.className = item.type === 'omml' ? 'sym template' : 'sym';
    btn.title = item.name;

    const glyph = document.createElement('span');
    glyph.textContent = item.symbol || item.preview || '?';
    btn.appendChild(glyph);

    if (item.type === 'omml') {
      const label = document.createElement('span');
      label.className = 'label';
      label.textContent = item.name;
      btn.appendChild(label);
    }

    const star = document.createElement('span');
    star.className = isFav(item) ? 'star on' : 'star';
    star.textContent = '★';
    star.title = 'Thêm/bỏ yêu thích';
    star.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleFav(item);
    });
    btn.appendChild(star);

    btn.addEventListener('click', () => insert(item));
    return btn;
  }

  function renderGrid(container, items, templates) {
    container.innerHTML = '';
    container.className = templates ? 'grid templates' : 'grid';
    items.forEach((item) => container.appendChild(buildButton(item)));
  }

  const toolTab = (name) => TOOL_TABS.find((t) => t.name === name);

  function tabButton(name) {
    const tab = document.createElement('button');
    tab.className = name === activeCategory ? 'tab tool active' : 'tab tool';
    tab.textContent = name;
    tab.addEventListener('click', () => {
      activeCategory = name;
      render();
    });
    return tab;
  }

  function renderTabs() {
    const tabs = document.getElementById('tabs');
    tabs.innerHTML = '';
    const row = document.createElement('div');
    row.className = 'tab-row';
    TOOL_TABS.forEach((t) => row.appendChild(tabButton(t.name)));
    tabs.appendChild(row);
  }

  // Bảng ký hiệu: ô tìm, ô chọn nhóm và lưới mẫu. Vẽ lại được riêng phần lưới nên
  // gõ tìm kiếm hay bấm ★ không làm mất con trỏ đang ở ô tìm.
  let symbolQuery = '';
  let refreshSymbolList = null;

  function renderSymbols(container) {
    const search = document.createElement('input');
    search.type = 'text';
    search.className = 'sym-search';
    search.placeholder = 'Tìm ký hiệu: alpha, phân số, nguyên hàm, | , ∫ ...';
    search.value = symbolQuery;

    const picker = document.createElement('select');
    picker.className = 'group-picker';
    SYMBOL_CATEGORIES.forEach((c) => {
      const opt = document.createElement('option');
      opt.value = c.name;
      opt.textContent = `${c.name} (${c.items.length})`;
      picker.appendChild(opt);
    });
    picker.value = activeGroup;

    const listBox = document.createElement('div');
    const hint = document.createElement('p');
    hint.className = 'hint';

    function paint() {
      const q = symbolQuery.trim().toLowerCase();
      picker.classList.toggle('hidden', !!q);
      const items = q
        ? allItems().filter(
            (i) =>
              i.name.toLowerCase().includes(q) ||
              (i.symbol || '').includes(q) ||
              (i.preview || '').toLowerCase().includes(q)
          )
        : SYMBOL_CATEGORIES.find((c) => c.name === activeGroup).items;
      const templates = items.some((i) => i.type === 'omml');
      hint.textContent = templates
        ? 'Công thức chèn ra có ô trống — bấm vào ô để gõ, Tab để nhảy sang ô kế. Đặt con trỏ trong một ô rồi bấm mẫu khác để lồng công thức vào nhau.'
        : '';
      hint.classList.toggle('hidden', !templates);
      if (items.length === 0) {
        listBox.className = '';
        listBox.innerHTML = '<div class="empty">Không tìm thấy ký hiệu nào.</div>';
        return;
      }
      renderGrid(listBox, items, templates);
    }

    search.addEventListener('input', () => {
      symbolQuery = search.value;
      paint();
    });
    picker.addEventListener('change', () => {
      activeGroup = picker.value;
      paint();
    });

    container.appendChild(search);
    container.appendChild(picker);
    container.appendChild(hint);
    container.appendChild(listBox);
    refreshSymbolList = paint;
    paint();
  }

  function render() {
    renderTabs();
    const panel = document.getElementById('panel');
    panel.innerHTML = '';
    panel.className = 'category tool-panel';
    refreshSymbolList = null;
    (toolTab(activeCategory) || TOOL_TABS[0]).render(panel);
  }

  // ------------------------------------------------ thanh nút nhanh (luôn thấy)

  // Thao tác lặp nhiều nhất khi soạn đề là "chèn câu tiếp theo" và "chèn lại
  // công thức vừa dùng", nên hai thứ đó nằm trên thanh dính, không phải cuộn tìm.
  const QUICK_KINDS = [
    ['mc', 'Trắc nghiệm'],
    ['tf', 'Đúng / Sai'],
    ['sa', 'Trả lời ngắn'],
    ['tl', 'Tự luận'],
  ];

  function nextNoBox() {
    return document.getElementById('quick-no');
  }

  function setNextNo(n) {
    const box = nextNoBox();
    if (box) box.value = String(n);
  }

  function insertNextQuestion() {
    const no = Math.max(1, parseInt((nextNoBox() || {}).value, 10) || 1);
    const kind = (document.getElementById('quick-kind') || {}).value || 'mc';
    const xml =
      kind === 'tl'
        ? essayBlock(no, 1, {
            diem: (document.getElementById('quick-diem') || {}).value || '2,0',
            subs: parseInt((document.getElementById('quick-subs') || {}).value, 10) || 0,
          })
        : questionBlock(kind, no, 1, parseInt((document.getElementById('quick-cols') || {}).value, 10) || 2);
    send(xml, true, `Đã chèn câu ${no}`);
    setNextNo(no + 1);
  }

  // Ô phụ đổi theo dạng câu đang chọn: trắc nghiệm cần số cột, tự luận cần điểm và số ý.
  function syncQuickExtras() {
    const kind = (document.getElementById('quick-kind') || {}).value || 'mc';
    const mc = document.getElementById('quick-cols');
    const tl = document.getElementById('quick-tl');
    if (mc) mc.classList.toggle('hidden', kind !== 'mc');
    if (tl) tl.classList.toggle('hidden', kind !== 'tl');
  }

  function renderQuickSyms() {
    const box = document.getElementById('quick-syms');
    if (!box) return;
    const byKey = new Map(allItems().map((i) => [keyOf(i), i]));
    const picks = favorites
      .concat(recent)
      .filter((k, i, arr) => arr.indexOf(k) === i)
      .map((k) => byKey.get(k))
      .filter(Boolean)
      .slice(0, 8);

    box.innerHTML = '';
    if (picks.length === 0) {
      box.appendChild(
        Object.assign(document.createElement('span'), {
          className: 'quick-empty',
          textContent: 'Ký hiệu vừa dùng và đánh dấu ★ sẽ hiện ở đây.',
        })
      );
      return;
    }
    picks.forEach((item) => {
      const btn = document.createElement('button');
      btn.className = 'quick-sym';
      btn.textContent = item.symbol || item.preview || '?';
      btn.title = item.name;
      btn.addEventListener('click', () => insert(item));
      box.appendChild(btn);
    });
  }

  function buildQuickBar() {
    const bar = document.createElement('div');
    bar.className = 'quickbar';
    bar.innerHTML =
      '<div class="quick-row">' +
      '<span class="quick-label">Câu</span>' +
      '<input id="quick-no" type="number" min="1" value="1" />' +
      '<select id="quick-kind"></select>' +
      '<select id="quick-cols">' +
      '<option value="2">2 cột</option><option value="1">1 cột</option><option value="4">4 cột</option>' +
      '</select>' +
      '<span id="quick-tl" class="quick-tl hidden">' +
      '<input id="quick-diem" type="text" value="2,0" title="Điểm mỗi câu" />' +
      '<input id="quick-subs" type="number" min="0" max="5" value="0" title="Số ý" />' +
      '</span>' +
      '<button id="quick-add" class="act primary">+ Chèn</button>' +
      '</div><div id="quick-syms" class="quick-syms"></div>';

    const kindSel = bar.querySelector('#quick-kind');
    QUICK_KINDS.forEach(([value, label]) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = label;
      kindSel.appendChild(opt);
    });
    kindSel.value = 'mc';
    kindSel.addEventListener('change', syncQuickExtras);
    bar.querySelector('#quick-add').addEventListener('click', insertNextQuestion);

    const app = document.getElementById('app');
    app.insertBefore(bar, document.getElementById('panel'));
    syncQuickExtras();
    renderQuickSyms();
  }

  function start(word) {
    inWord = word;
    // Các tab công cụ ở tools.js chèn nội dung qua đây.
    window.Pane = {
      inWord: word,
      toast,
      insertOoxml: (xml, label) => send(xml, true, label),
      renderSymbols,
    };
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('app').classList.remove('hidden');
    // Ô tìm và mục yêu thích cũ trong HTML đã được thay bằng bảng ký hiệu ở tab
    // Chèn và thanh nút nhanh, nên ẩn hẳn.
    document.querySelector('.search-bar').classList.add('hidden');
    document.getElementById('favorites-section').classList.add('hidden');
    buildQuickBar();
    render();
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => start(info.host === Office.HostType.Word));
  } else {
    start(false);
  }
})();
