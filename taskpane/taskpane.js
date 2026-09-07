(function () {
  const FAV_KEY = 'mathSymbols.favorites';
  let activeCategory = TOOL_TABS[0].name;
  let favorites = loadFavorites();
  let inWord = false;

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

  // Công cụ soạn đề là phần chính, nên đứng trước; bảng ký hiệu/công thức là
  // phần phụ, gom xuống dưới một nhãn nhỏ cho đỡ chiếm chỗ.
  function renderTabs() {
    const tabs = document.getElementById('tabs');
    tabs.innerHTML = '';

    const tools = document.createElement('div');
    tools.className = 'tab-row';
    TOOL_TABS.forEach((t) => tools.appendChild(tabButton(t.name, 'tab tool')));
    tabs.appendChild(tools);

    const label = document.createElement('div');
    label.className = 'tab-group-label';
    label.textContent = 'Ký hiệu & công thức';
    tabs.appendChild(label);

    const symbols = document.createElement('div');
    symbols.className = 'tab-row';
    SYMBOL_CATEGORIES.forEach((c) => symbols.appendChild(tabButton(c.name, 'tab')));
    tabs.appendChild(symbols);
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

    const favItems = allItems().filter(isFav);
    if (favItems.length > 0) {
      favSection.classList.remove('hidden');
      renderGrid(document.getElementById('favorites-grid'), favItems, false);
    } else {
      favSection.classList.add('hidden');
    }

    const category = SYMBOL_CATEGORIES.find((c) => c.name === activeCategory);
    const isTemplates = category.items.some((i) => i.type === 'omml');
    panel.innerHTML = `<h3>${category.name}</h3>`;
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
    render();
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => start(info.host === Office.HostType.Word));
  } else {
    start(false);
  }
})();
