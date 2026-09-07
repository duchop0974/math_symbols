(function () {
  const FAV_KEY = 'mathSymbols.favorites';
  const RECENT_KEY = 'mathSymbols.recent';
  const SYMBOL_TAB = 'Ký hiệu';

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
    render();
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

  function tabButton(name, kind) {
    const tab = document.createElement('button');
    tab.className = name === activeCategory ? `${kind} active` : kind;
    tab.textContent = name;
    tab.addEventListener('click', () => {
      activeCategory = name;
      document.getElementById('search').value = '';
      render();
    });
    return tab;
  }

  // Bốn tab trên một hàng: ba công cụ soạn đề và một tab gom toàn bộ ký hiệu.
  function renderTabs() {
    const tabs = document.getElementById('tabs');
    tabs.innerHTML = '';
    const row = document.createElement('div');
    row.className = 'tab-row';
    TOOL_TABS.forEach((t) => row.appendChild(tabButton(t.name, 'tab tool')));
    row.appendChild(tabButton(SYMBOL_TAB, 'tab'));
    tabs.appendChild(row);
  }

  // Nhóm ký hiệu giờ là ô chọn trong tab, không còn mỗi nhóm một tab.
  function renderGroupPicker(panel) {
    const picker = document.createElement('select');
    picker.className = 'group-picker';
    SYMBOL_CATEGORIES.forEach((c) => {
      const opt = document.createElement('option');
      opt.value = c.name;
      opt.textContent = `${c.name} (${c.items.length})`;
      picker.appendChild(opt);
    });
    picker.value = activeGroup;
    picker.addEventListener('change', (e) => {
      activeGroup = e.target.value;
      render();
    });
    panel.appendChild(picker);
  }

  function render() {
    const query = document.getElementById('search').value.trim().toLowerCase();
    const panel = document.getElementById('panel');
    const favSection = document.getElementById('favorites-section');

    renderTabs();

    if (query) {
      favSection.classList.add('hidden');
      panel.className = 'category';
      // Soi cả `preview` để gõ thẳng ký hiệu (vd "|", "∫", "√") cũng tìm ra mẫu.
      const matches = allItems().filter(
        (i) =>
          i.name.toLowerCase().includes(query) ||
          (i.symbol || '').includes(query) ||
          (i.preview || '').toLowerCase().includes(query)
      );
      panel.innerHTML = '<h3>Kết quả tìm kiếm</h3>';
      const grid = document.createElement('div');
      panel.appendChild(grid);
      if (matches.length === 0) {
        panel.innerHTML = '<div class="empty">Không tìm thấy ký hiệu nào.</div>';
      } else {
        renderGrid(grid, matches, matches.some((m) => m.type === 'omml'));
      }
      return;
    }

    // Ô tìm kiếm chỉ tra được ký hiệu, nên giấu đi khi đang ở tab công cụ.
    const searchBar = document.querySelector('.search-bar');
    const tool = toolTab(activeCategory);
    if (tool) {
      searchBar.classList.add('hidden');
      favSection.classList.add('hidden');
      panel.innerHTML = '';
      panel.className = 'category tool-panel';
      tool.render(panel);
      return;
    }
    searchBar.classList.remove('hidden');
    panel.className = 'category';

    // Mục yêu thích cũ đã được thanh nút nhanh thay thế, giấu đi cho gọn.
    favSection.classList.add('hidden');

    const category = SYMBOL_CATEGORIES.find((c) => c.name === activeGroup);
    const isTemplates = category.items.some((i) => i.type === 'omml');
    panel.innerHTML = '';
    renderGroupPicker(panel);
    if (isTemplates) {
      const hint = document.createElement('p');
      hint.className = 'hint';
      hint.textContent =
        'Công thức chèn ra có ô trống — bấm vào ô để gõ số, nhấn Tab để nhảy sang ô kế. Đặt con trỏ trong một ô rồi bấm mẫu khác để lồng công thức vào nhau.';
      panel.appendChild(hint);
    }
    const grid = document.createElement('div');
    panel.appendChild(grid);
    renderGrid(grid, category.items, isTemplates);
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

  // Ô "Bắt đầu từ câu" của tab Đề thi và ô trên thanh nhanh phải luôn khớp nhau.
  function setNextNo(n) {
    const box = nextNoBox();
    if (box) box.value = String(n);
    const panelBox = document.getElementById('q-start');
    if (panelBox) panelBox.value = String(n);
  }

  function insertNextQuestion() {
    const no = Math.max(1, parseInt((nextNoBox() || {}).value, 10) || 1);
    const kind = (document.getElementById('quick-kind') || {}).value || 'mc';
    const cols = parseInt((document.getElementById('skel-cols') || {}).value, 10) || 2;
    const xml =
      kind === 'tl'
        ? essayBlock(no, 1, {
            diem: (document.getElementById('skel-diem') || {}).value || '2,0',
            subs: parseInt((document.getElementById('skel-subs') || {}).value, 10) || 0,
          })
        : questionBlock(kind, no, 1, cols);
    send(xml, true, `Đã chèn câu ${no}`);
    setNextNo(no + 1);
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
    bar.querySelector('#quick-add').addEventListener('click', insertNextQuestion);

    const app = document.getElementById('app');
    app.insertBefore(bar, document.getElementById('panel'));
    renderQuickSyms();
  }

  function start(word) {
    inWord = word;
    // Các tab công cụ ở tools.js chèn nội dung qua đây.
    window.Pane = {
      inWord: word,
      toast,
      insertOoxml: (xml, label) => send(xml, true, label),
    };
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('app').classList.remove('hidden');
    document.getElementById('search').addEventListener('input', render);
    buildQuickBar();
    render();
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => start(info.host === Office.HostType.Word));
  } else {
    start(false);
  }
})();
