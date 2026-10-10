// Таблица с фильтрами в заголовках, как в Excel. Работает офлайн, без внешних библиотек.
//
// Python (app/ui_table.py) передаёт в data: колонки, строки (уже отфильтрованные),
// текущие фильтры и сортировку. Когда пользователь нажимает «ОК» в окне фильтра или
// щёлкает по заголовку, таблица отправляет в Python новое состояние view =
// {filters: {...}, sort: {code, desc} | null}. Фильтрует Python, сортирует таблица сама
// (и сообщает сортировку в Python, чтобы она попала в выгрузку).

const ROW_H = 30;
const HEAD_H = 40;
const MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];

const nfMoney = new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 2 });
const nfPercent = new Intl.NumberFormat("ru-RU", { minimumFractionDigits: 1, maximumFractionDigits: 1 });

export default function (component) {
  const { data, setStateValue, parentElement } = component;
  let table = parentElement.__prognozTable;
  if (!table || !table.root.isConnected) {
    table = new Table(parentElement);
    parentElement.__prognozTable = table;
  }
  table.send = (view) => setStateValue("view", view);
  table.update(data);
}

// ---------- форматирование ----------

function formatCell(col, v) {
  if (v === null || v === undefined) return "";
  switch (col.fmt) {
    case "money": return nfMoney.format(v);
    case "percent": return nfPercent.format(v * 100) + " %";
    case "date": return v.slice(8, 10) + "." + v.slice(5, 7) + "." + v.slice(0, 4);
    default: return String(v);
  }
}

function formatOption(col, v) {
  return v === null ? "(пусто)" : formatCell(col, v);
}

function isNumeric(col) {
  return col.fmt === "money" || col.fmt === "percent" || col.fmt === "int";
}

function compareValues(a, b) {
  if (a === b) return 0;
  if (a === null || a === undefined) return 1;   // пустые всегда в конце
  if (b === null || b === undefined) return -1;
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a).localeCompare(String(b), "ru", { numeric: true });
}

function parseNumber(text) {
  const clean = String(text).replace(/[\s ]/g, "").replace(",", ".");
  if (clean === "") return null;
  const n = Number(clean);
  return Number.isFinite(n) ? n : NaN;
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

const FUNNEL = '<svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M3 4h18l-7 8.5V19l-4 2v-8.5z"/></svg>';

// ---------- таблица ----------

class Table {
  constructor(parent) {
    this.root = el("div", "pt");
    this.bar = el("div", "pt-bar");
    this.count = el("span", "pt-count");
    this.resetBtn = el("button", "pt-btn", "Сбросить все фильтры");
    this.fsBtn = el("button", "pt-btn", "⛶ На весь экран");
    this.bar.append(this.count, this.resetBtn, this.fsBtn);

    this.scroll = el("div", "pt-scroll");
    this.inner = el("div", "pt-inner");
    this.head = el("div", "pt-head");
    this.body = el("div", "pt-body");
    this.inner.append(this.head, this.body);
    this.scroll.append(this.inner);
    this.root.append(this.bar, this.scroll);
    parent.append(this.root);

    this.dataId = null;
    this.columns = [];
    this.rows = [];
    this.order = [];
    this.widths = {};
    this.filters = {};
    this.sort = null;

    this.scroll.addEventListener("scroll", () => this.renderRows());
    window.addEventListener("resize", () => this.renderRows());
    this.resetBtn.onclick = () => this.send({ filters: {}, sort: this.sort });
    this.fsBtn.onclick = () => this.toggleFullscreen();
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      if (this.popup) this.closeFilter();
      else if (this.root.classList.contains("pt-max")) this.toggleFullscreen();
    });
  }

  update(data) {
    this.root.style.height = data.height + "px";
    this.filters = data.filters || {};
    this.sort = data.sort || null;
    this.columns = data.columns;
    if (data.data_id !== this.dataId) {
      this.dataId = data.data_id;
      this.rows = data.rows;
      this.scroll.scrollTop = 0;
    }
    for (const col of this.columns) {
      if (!this.widths[col.code]) this.widths[col.code] = this.guessWidth(col);
    }
    this.count.textContent = data.total === this.rows.length
      ? `Строк: ${this.rows.length}`
      : `Показано строк: ${this.rows.length} из ${data.total}`;
    if (data.truncated) this.count.textContent += ` (в таблице первые ${this.rows.length})`;
    this.resetBtn.disabled = Object.keys(this.filters).length === 0;
    this.applySort();
    this.renderHead();
    this.renderRows();
  }

  guessWidth(col) {
    let longest = col.label.length * 7 + 50;
    const sample = Math.min(this.rows.length, 200);
    for (let i = 0; i < sample; i++) {
      const text = formatCell(col, this.rows[i][col.index]);
      longest = Math.max(longest, text.length * 7.5 + 20);
    }
    return Math.round(Math.min(Math.max(longest, 70), 320));
  }

  applySort() {
    const n = this.rows.length;
    this.order = Array.from({ length: n }, (_, i) => i);
    if (!this.sort) return;
    const col = this.columns.find((c) => c.code === this.sort.code);
    if (!col) return;
    const k = col.index;
    const sign = this.sort.desc ? -1 : 1;
    const rows = this.rows;
    this.order.sort((a, b) => {
      const va = rows[a][k], vb = rows[b][k];
      // Пустые остаются внизу при любом направлении.
      if (va === null && vb === null) return a - b;
      if (va === null) return 1;
      if (vb === null) return -1;
      return sign * compareValues(va, vb) || a - b;
    });
  }

  totalWidth() {
    return this.columns.reduce((s, c) => s + this.widths[c.code], 0);
  }

  renderHead() {
    this.head.innerHTML = "";
    this.inner.style.width = this.totalWidth() + "px";
    for (const col of this.columns) {
      const th = el("div", "pt-th");
      th.style.width = this.widths[col.code] + "px";
      const title = el("span", "pt-title", col.label);
      title.title = col.label + " — щёлкните, чтобы отсортировать";
      title.onclick = () => this.cycleSort(col);
      const sortMark = el("span", "pt-sort");
      if (this.sort && this.sort.code === col.code) sortMark.textContent = this.sort.desc ? "▼" : "▲";
      th.append(title, sortMark);
      if (col.kind) {
        const btn = el("button", "pt-fbtn");
        const active = col.code in this.filters;
        if (active) { btn.classList.add("active"); btn.innerHTML = FUNNEL; } else { btn.textContent = "▾"; }
        btn.title = active ? "Фильтр включён" : "Фильтр";
        btn.onclick = (e) => { e.stopPropagation(); this.openFilter(col, th); };
        th.append(btn);
      }
      const grip = el("div", "pt-resize");
      grip.onmousedown = (e) => this.startResize(e, col);
      th.append(grip);
      this.head.append(th);
    }
  }

  renderRows() {
    const n = this.order.length;
    this.body.style.height = Math.max(n * ROW_H, 1) + "px";
    this.body.innerHTML = "";
    if (n === 0) {
      this.body.append(el("div", "pt-empty", "Нет строк, подходящих под фильтры."));
      this.body.style.height = "auto";
      return;
    }
    const top = Math.max(0, this.scroll.scrollTop - HEAD_H);
    const first = Math.max(0, Math.floor(top / ROW_H) - 5);
    const visible = Math.ceil(this.scroll.clientHeight / ROW_H) + 10;
    const last = Math.min(n, first + visible);
    const width = this.totalWidth() + "px";
    const frag = document.createDocumentFragment();
    for (let i = first; i < last; i++) {
      const row = this.rows[this.order[i]];
      const tr = el("div", "pt-tr");
      tr.style.top = i * ROW_H + "px";
      tr.style.width = width;
      for (const col of this.columns) {
        const text = formatCell(col, row[col.index]);
        const td = el("div", isNumeric(col) ? "pt-td num" : "pt-td", text);
        td.style.width = this.widths[col.code] + "px";
        if (text.length > 12) td.title = text;
        tr.append(td);
      }
      frag.append(tr);
    }
    this.body.append(frag);
  }

  cycleSort(col) {
    let sort;
    if (!this.sort || this.sort.code !== col.code) sort = { code: col.code, desc: false };
    else if (!this.sort.desc) sort = { code: col.code, desc: true };
    else sort = null;
    this.setSort(sort);
  }

  setSort(sort) {
    this.sort = sort;
    this.applySort();
    this.renderHead();
    this.renderRows();
    this.send({ filters: this.filters, sort });
  }

  startResize(e, col) {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = this.widths[col.code];
    const move = (ev) => {
      this.widths[col.code] = Math.max(50, startW + ev.clientX - startX);
      this.renderHead();
      this.renderRows();
    };
    const up = () => {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
    };
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
  }

  // «На весь экран» — таблица разворачивается поверх страницы (как у стандартной таблицы
  // Streamlit). Настоящий полноэкранный режим браузера не используем: его могут запрещать.
  toggleFullscreen() {
    const on = this.root.classList.toggle("pt-max");
    this.fsBtn.textContent = on ? "✕ Свернуть (Esc)" : "⛶ На весь экран";
    this.renderRows();
  }

  // ---------- окно фильтра ----------

  openFilter(col, th) {
    this.closeFilter();
    const overlay = el("div", "pt-overlay");
    overlay.onclick = () => this.closeFilter();
    const pop = el("div", "pt-pop");
    pop.onclick = (e) => e.stopPropagation();
    this.root.append(overlay, pop);
    this.popup = { overlay, pop };

    pop.append(el("div", "pt-pop-title", col.label));
    const sortRow = el("div", "pt-pop-sort");
    const asc = el("button", "pt-btn", col.kind === "values" ? "↑ А → Я" : "↑ По возрастанию");
    const desc = el("button", "pt-btn", col.kind === "values" ? "↓ Я → А" : "↓ По убыванию");
    if (isNumeric(col) || col.fmt === "date") { asc.textContent = "↑ По возрастанию"; desc.textContent = "↓ По убыванию"; }
    asc.onclick = () => { this.closeFilter(); this.setSort({ code: col.code, desc: false }); };
    desc.onclick = () => { this.closeFilter(); this.setSort({ code: col.code, desc: true }); };
    sortRow.append(asc, desc);
    pop.append(sortRow);

    const current = this.filters[col.code] || null;
    let editor;
    if (col.kind === "values") editor = valuesEditor(col, current);
    else if (col.kind === "months") editor = monthsEditor(col, current);
    else if (col.kind === "range") editor = rangeEditor(col, current);
    else editor = containsEditor(col, current);
    pop.append(...editor.nodes);

    const buttons = el("div", "pt-pop-buttons");
    const clear = el("button", "pt-btn pt-clear", "Удалить фильтр");
    clear.disabled = !current;
    clear.onclick = () => this.applyFilter(col, null);
    const ok = el("button", "pt-btn primary", "ОК");
    ok.onclick = () => {
      const result = editor.result();
      if (result === undefined) return;  // ввод с ошибкой, редактор сам показал подсказку
      this.applyFilter(col, result);
    };
    const cancel = el("button", "pt-btn", "Отмена");
    cancel.onclick = () => this.closeFilter();
    buttons.append(clear, ok, cancel);
    pop.append(buttons);
    if (editor.onValidity) editor.onValidity((valid) => { ok.disabled = !valid; });

    // Окно под заголовком, но в пределах экрана.
    const r = th.getBoundingClientRect();
    const w = 300;
    const left = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
    pop.style.left = left + "px";
    const below = r.bottom + 4;
    const maxH = Math.min(460, window.innerHeight - 16);
    pop.style.top = Math.max(8, Math.min(below, window.innerHeight - maxH - 8)) + "px";
    pop.style.maxHeight = maxH + "px";
    if (editor.focus) editor.focus();
  }

  closeFilter() {
    if (!this.popup) return;
    this.popup.overlay.remove();
    this.popup.pop.remove();
    this.popup = null;
  }

  applyFilter(col, filter) {
    this.closeFilter();
    const filters = { ...this.filters };
    if (filter) filters[col.code] = filter; else delete filters[col.code];
    this.filters = filters;
    this.renderHead();
    this.send({ filters, sort: this.sort });
  }
}

// ---------- редакторы фильтров ----------
// Каждый возвращает { nodes, result(), focus?, onValidity? }.
// result(): объект фильтра, null (фильтр не нужен — всё отмечено) или undefined (ошибка ввода).

function checkItem(label, checked, cls) {
  const item = el("label", "pt-item" + (cls ? " " + cls : ""));
  const box = el("input");
  box.type = "checkbox";
  box.checked = checked;
  item.append(box, el("span", "pt-label", label));
  return { item, box };
}

function valuesEditor(col, current) {
  const options = col.options;  // сырые значения, null = пустые ячейки
  const selected = new Set(current ? current.values.map((v) => JSON.stringify(v)) : options.map((v) => JSON.stringify(v)));
  const search = el("input", "pt-input");
  search.placeholder = "Поиск";
  const list = el("div", "pt-list");
  let validity = () => {};

  const items = options.map((v) => {
    const key = JSON.stringify(v);
    const { item, box } = checkItem(formatOption(col, v), selected.has(key));
    box.onchange = () => { if (box.checked) selected.add(key); else selected.delete(key); syncAll(); };
    return { v, key, item, box, text: formatOption(col, v).toLowerCase() };
  });
  const all = checkItem("(Выделить все)", true);
  all.box.onchange = () => {
    for (const it of visibleItems()) {
      it.box.checked = all.box.checked;
      if (all.box.checked) selected.add(it.key); else selected.delete(it.key);
    }
    syncAll();
  };

  function visibleItems() {
    const q = search.value.trim().toLowerCase();
    return items.filter((it) => !q || it.text.includes(q));
  }
  function syncAll() {
    const vis = visibleItems();
    const on = vis.filter((it) => it.box.checked).length;
    all.box.checked = vis.length > 0 && on === vis.length;
    all.box.indeterminate = on > 0 && on < vis.length;
    validity(search.value.trim() ? on > 0 : selected.size > 0);
  }
  function render() {
    list.innerHTML = "";
    list.append(all.item);
    for (const it of visibleItems()) list.append(it.item);
    syncAll();
  }
  search.oninput = render;
  render();

  return {
    nodes: [search, list],
    focus: () => search.focus(),
    onValidity: (fn) => { validity = fn; syncAll(); },
    result: () => {
      // Как в Excel: если в поиске что-то введено, остаются только найденные и отмеченные.
      const pool = search.value.trim() ? visibleItems() : items;
      const chosen = pool.filter((it) => selected.has(it.key)).map((it) => it.v);
      if (chosen.length === 0) return undefined;
      if (chosen.length === options.length) return null;
      return { kind: "values", values: chosen };
    },
  };
}

function monthsEditor(col, current) {
  // Значения фильтра — строки "2026-01"; null — пустые ячейки.
  const allKeys = [];
  for (const y of col.options.years) for (const m of y.months) allKeys.push(`${y.year}-${String(m).padStart(2, "0")}`);
  if (col.options.has_empty) allKeys.push(null);
  const selected = new Set(current ? current.values : allKeys);
  const list = el("div", "pt-list");
  let validity = () => {};
  const yearRows = [];

  const all = checkItem("(Выделить все)", true);
  all.box.onchange = () => {
    selected.clear();
    if (all.box.checked) allKeys.forEach((k) => selected.add(k));
    sync();
  };
  list.append(all.item);

  for (const y of col.options.years) {
    const keys = y.months.map((m) => `${y.year}-${String(m).padStart(2, "0")}`);
    const yr = checkItem(String(y.year), true);
    const toggle = el("span", "pt-toggle", "▸");
    yr.item.prepend(toggle);
    const children = el("div");
    children.style.display = "none";
    const monthBoxes = y.months.map((m, i) => {
      const { item, box } = checkItem(MONTHS[m - 1], true, "child");
      box.onchange = () => { if (box.checked) selected.add(keys[i]); else selected.delete(keys[i]); sync(); };
      children.append(item);
      return box;
    });
    toggle.onclick = (e) => {
      e.preventDefault();
      const open = children.style.display === "none";
      children.style.display = open ? "block" : "none";
      toggle.textContent = open ? "▾" : "▸";
    };
    yr.box.onchange = () => {
      keys.forEach((k) => (yr.box.checked ? selected.add(k) : selected.delete(k)));
      sync();
    };
    list.append(yr.item, children);
    yearRows.push({ keys, box: yr.box, monthBoxes });
  }
  let emptyBox = null;
  if (col.options.has_empty) {
    const empty = checkItem("(пусто)", true);
    empty.box.onchange = () => { if (empty.box.checked) selected.add(null); else selected.delete(null); sync(); };
    list.append(empty.item);
    emptyBox = empty.box;
  }

  function sync() {
    for (const yr of yearRows) {
      const on = yr.keys.filter((k) => selected.has(k)).length;
      yr.box.checked = on === yr.keys.length;
      yr.box.indeterminate = on > 0 && on < yr.keys.length;
      yr.monthBoxes.forEach((b, i) => { b.checked = selected.has(yr.keys[i]); });
    }
    if (emptyBox) emptyBox.checked = selected.has(null);
    all.box.checked = selected.size === allKeys.length;
    all.box.indeterminate = selected.size > 0 && selected.size < allKeys.length;
    validity(selected.size > 0);
  }
  sync();

  return {
    nodes: [list],
    onValidity: (fn) => { validity = fn; sync(); },
    result: () => {
      if (selected.size === 0) return undefined;
      if (selected.size === allKeys.length) return null;
      return { kind: "months", values: allKeys.filter((k) => selected.has(k)) };
    },
  };
}

function rangeEditor(col, current) {
  const scale = col.fmt === "percent" ? 100 : 1;
  const unit = col.fmt === "percent" ? ", %" : "";
  const box = el("div", "pt-range");
  const low = el("input", "pt-input");
  const high = el("input", "pt-input");
  const show = (v) => (v === null || v === undefined ? "" : nfMoney.format(v * scale));
  low.value = current ? show(current.low) : "";
  high.value = current ? show(current.high) : "";
  low.placeholder = "мин. " + show(col.options.min);
  high.placeholder = "макс. " + show(col.options.max);
  box.append(el("span", null, "от" + unit), low, el("span", null, "до" + unit), high);
  const note = el("div", "pt-note", "Пустое поле — без ограничения. Строки с пустой ячейкой при фильтре не показываются.");

  return {
    nodes: [box, note],
    focus: () => low.focus(),
    result: () => {
      const a = parseNumber(low.value), b = parseNumber(high.value);
      if (Number.isNaN(a) || Number.isNaN(b)) { note.textContent = "Введите число, например 1 000 000 или 12,5."; return undefined; }
      if (a === null && b === null) return null;
      return { kind: "range", low: a === null ? null : a / scale, high: b === null ? null : b / scale };
    },
  };
}

function containsEditor(col, current) {
  const input = el("input", "pt-input");
  input.placeholder = "Текст для поиска";
  input.value = current ? current.text : "";
  const note = el("div", "pt-note", "Разных значений слишком много для списка, поэтому поиск по тексту.");
  return {
    nodes: [input, note],
    focus: () => input.focus(),
    result: () => (input.value.trim() ? { kind: "contains", text: input.value.trim() } : null),
  };
}
