(function () {
  const FAV_KEY = 'mathSymbols.favorites';
  let activeCategory = SYMBOL_CATEGORIES[0].name;
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

  function insert(item) {
    const isOmml = item.type === 'omml';
    const data = isOmml ? item.ooxml : item.symbol;

    if (!inWord) {
      navigator.clipboard?.writeText(data);
      toast(`Chế độ xem thử: đã copy "${item.name}" vào clipboard`);
      return;
    }

    Office.context.document.setSelectedDataAsync(
      data,
      { coercionType: isOmml ? Office.CoercionType.Ooxml : Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          toast(`Lỗi chèn: ${result.error.message}`, true);
        } else {
          toast(`Đã chèn ${item.symbol || item.name}`);
        }
      }
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

  function renderTabs() {
    const tabs = document.getElementById('tabs');
    tabs.innerHTML = '';
    SYMBOL_CATEGORIES.forEach((cat) => {
      const tab = document.createElement('button');
      tab.className = cat.name === activeCategory ? 'tab active' : 'tab';
      tab.textContent = cat.name;
      tab.addEventListener('click', () => {
        activeCategory = cat.name;
        document.getElementById('search').value = '';
        render();
      });
      tabs.appendChild(tab);
    });
  }

  function render() {
    const query = document.getElementById('search').value.trim().toLowerCase();
    const panel = document.getElementById('panel');
    const favSection = document.getElementById('favorites-section');

    renderTabs();

    if (query) {
      favSection.classList.add('hidden');
      const matches = allItems().filter(
        (i) => i.name.toLowerCase().includes(query) || (i.symbol || '').includes(query)
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

    const favItems = allItems().filter(isFav);
    if (favItems.length > 0) {
      favSection.classList.remove('hidden');
      renderGrid(document.getElementById('favorites-grid'), favItems, false);
    } else {
      favSection.classList.add('hidden');
    }

    const category = SYMBOL_CATEGORIES.find((c) => c.name === activeCategory);
    panel.innerHTML = `<h3>${category.name}</h3>`;
    const grid = document.createElement('div');
    panel.appendChild(grid);
    renderGrid(grid, category.items, category.items.some((i) => i.type === 'omml'));
  }

  function start(word) {
    inWord = word;
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
