const state = {
  data: null,
  students: [],
  hws: [],
  hwCount: 0,
  countsByAccepted: {},
  statuses: new Set(["хорошие показатели", "средние показатели", "низкие показатели"]),
  minAccepted: 0,
  topN: 0,
  rankMode: "position",
  colorMode: "status",
  distMode: "filtered",
  highlightOn: true,
  selected: null,
  compare: new Set(),
  plotBound: false,
};

const el = {
  q: document.getElementById("q"),
  suggest: document.getElementById("suggest"),
  compareInput: document.getElementById("compareInput"),
  compareSuggest: document.getElementById("compareSuggest"),
  minAcc: document.getElementById("minAcc"),
  minAccLabel: document.getElementById("minAccLabel"),
  minAccTotal: document.getElementById("minAccTotal"),
  topN: document.getElementById("topN"),
  rankMode: document.getElementById("rankMode"),
  colorMode: document.getElementById("colorMode"),
  distMode: document.getElementById("distMode"),
  btnTheme: document.getElementById("btnTheme"),
  btnReset: document.getElementById("btnReset"),
  btnExport: document.getElementById("btnExport"),
  btnRefresh: document.getElementById("btnRefresh"),
  btnHighlight: document.getElementById("btnHighlight"),
  btnAddCompare: document.getElementById("btnAddCompare"),
  btnClearCompare: document.getElementById("btnClearCompare"),
  plotScatter: document.getElementById("plotScatter"),
  plotDist: document.getElementById("plotDist"),
  plotCompare: document.getElementById("plotCompare"),
  detail: document.getElementById("detail"),
  compareList: document.getElementById("compareList"),
  selPill: document.getElementById("selPill"),
  metaStudents: document.getElementById("metaStudents"),
  metaFiles: document.getElementById("metaFiles"),
  chartMeta: document.getElementById("chartMeta"),
  titleRange: document.getElementById("titleRange"),
  chartRange: document.getElementById("chartRange"),
  kTotal: document.getElementById("kTotal"),
  kFiltered: document.getElementById("kFiltered"),
  kFilteredSub: document.getElementById("kFilteredSub"),
  kMean: document.getElementById("kMean"),
  kGoodShare: document.getElementById("kGoodShare"),
};

const LS_COMPARE = "score_compare_v1";
const LS_THEME = "score_theme_v1";

function escHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function getCss(varName) {
  return getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
}

function debounce(fn, ms = 160) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  const size = clean.length === 3 ? 1 : 2;
  const nums = [];
  for (let i = 0; i < clean.length; i += size) {
    const chunk = clean.slice(i, i + size);
    nums.push(parseInt(size === 1 ? chunk + chunk : chunk, 16));
  }
  return { r: nums[0] || 0, g: nums[1] || 0, b: nums[2] || 0 };
}

function mixColors(a, b, t) {
  const ca = hexToRgb(a);
  const cb = hexToRgb(b);
  const mix = (x, y) => Math.round(x + (y - x) * t);
  return `rgb(${mix(ca.r, cb.r)}, ${mix(ca.g, cb.g)}, ${mix(ca.b, cb.b)})`;
}

function setTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem(LS_THEME, theme);
  el.btnTheme.textContent = theme === "dark" ? "Тема: темная" : "Тема: светлая";
  render();
}

function loadTheme() {
  const saved = localStorage.getItem(LS_THEME);
  if (saved) {
    setTheme(saved);
  } else {
    el.btnTheme.textContent = "Тема: светлая";
  }
}

function loadCompare() {
  try {
    const raw = localStorage.getItem(LS_COMPARE);
    if (!raw) return;
    const names = JSON.parse(raw);
    if (Array.isArray(names)) {
      names.forEach((nm) => state.compare.add(nm));
    }
  } catch (err) {
    console.warn("compare load failed", err);
  }
}

function saveCompare() {
  localStorage.setItem(LS_COMPARE, JSON.stringify([...state.compare]));
}

function formatHwRange() {
  if (!state.hws.length) return "HW";
  const first = state.hws[0]?.label || "HW";
  const last = state.hws[state.hws.length - 1]?.label || first;
  return state.hws.length === 1 ? first : `${first}–${last}`;
}

function updateRangeLabels() {
  const range = formatHwRange();
  if (el.titleRange) {
    el.titleRange.textContent = `Рейтинг ${range}`;
  }
  if (el.chartRange) {
    el.chartRange.textContent = range;
  }
  document.title = `Institute Score Dashboard — ${range}`;
}

async function fetchData(force = false) {
  const endpoint = force ? "/api/refresh" : "/api/data";
  const options = force ? { method: "POST" } : {};
  const res = await fetch(endpoint, options);
  if (!res.ok) {
    throw new Error("data load failed");
  }
  const data = await res.json();
  applyData(data);
}

function applyData(data) {
  state.data = data;
  state.students = data.students || [];
  state.hws = data.hws || [];
  state.hwCount = data.meta?.hw_count ?? 0;
  state.countsByAccepted = data.stats?.counts_by_accepted || {};

  el.metaStudents.textContent = `студентов: ${state.students.length}`;
  el.metaFiles.textContent = `файлов: ${state.hws.length}`;
  el.chartMeta.textContent = `обновлено: ${data.meta?.generated_at ?? "—"}`;
  updateRangeLabels();

  el.minAcc.max = String(state.hwCount);
  el.minAccTotal.textContent = String(state.hwCount || 0);
  if (state.minAccepted > state.hwCount) {
    state.minAccepted = state.hwCount;
    el.minAcc.value = String(state.minAccepted);
    el.minAccLabel.textContent = String(state.minAccepted);
  }

  const validNames = new Set(state.students.map((s) => s.name));
  [...state.compare].forEach((nm) => {
    if (!validNames.has(nm)) state.compare.delete(nm);
  });
  saveCompare();

  render();
}

function applyFilters() {
  let arr = state.students.filter(
    (s) => state.statuses.has(s.status) && s.accepted >= state.minAccepted
  );
  if (state.topN && Number(state.topN) > 0) {
    arr = arr.filter((s) => s.rank <= Number(state.topN));
  }
  return arr;
}

function computeStats(arr) {
  const n = arr.length;
  const mean = n ? arr.reduce((a, b) => a + b.accepted, 0) / n : 0;
  const good = arr.filter((s) => s.status === "хорошие показатели").length;
  const goodShare = n ? good / n : 0;
  return { n, mean, goodShare };
}

function computeRankValue(arr, mode) {
  if (mode === "position") {
    return arr.map((s) => s.rank);
  }
  if (mode === "dense") {
    return arr.map((s) => (state.hwCount + 1) - s.accepted);
  }
  const counts = {};
  Object.entries(state.countsByAccepted).forEach(([k, v]) => {
    counts[Number(k)] = v;
  });
  const higherCounts = {};
  for (let a = 0; a <= state.hwCount; a++) {
    let sum = 0;
    for (let b = a + 1; b <= state.hwCount; b++) {
      sum += counts[b] || 0;
    }
    higherCounts[a] = sum;
  }
  return arr.map((s) => 1 + higherCounts[s.accepted]);
}

function colorForPoint(student) {
  if (state.colorMode === "status") {
    if (student.status === "хорошие показатели") return getCss("--good");
    if (student.status === "средние показатели") return getCss("--mid");
    if (student.status === "низкие показатели") return getCss("--bad");
    return getCss("--bad");
  }
  const t = state.hwCount ? student.accepted / state.hwCount : 0;
  return mixColors(getCss("--bad"), getCss("--good"), t);
}

function buildScatterTraces(arr) {
  const xVals = computeRankValue(arr, state.rankMode);
  const groups = new Map();
  for (let i = 0; i < arr.length; i++) {
    const s = arr[i];
    const key = s.accepted;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ s, x: xVals[i] });
  }

  const traces = [];
  for (let a = state.hwCount; a >= 0; a--) {
    const pts = groups.get(a) || [];
    traces.push({
      type: "scattergl",
      mode: "markers",
      name: `${a}/${state.hwCount}`,
      x: pts.map((p) => p.x),
      y: pts.map((p) => p.s.accepted),
      customdata: pts.map((p) => [
        p.s.name,
        p.s.percent,
        p.s.status,
        p.s.rank,
        p.s.accepted,
      ]),
      hovertemplate:
        "<b>%{customdata[0]}</b><br>" +
        "Позиция: %{customdata[3]}<br>" +
        "Зачтено: %{customdata[4]}/" +
        state.hwCount +
        " (%{customdata[1]}%)<br>" +
        "Статус: %{customdata[2]}<extra></extra>",
      marker: {
        size: 7,
        opacity: 0.9,
        color: pts.map((p) => colorForPoint(p.s)),
        line: { width: 0 },
      },
      showlegend: true,
    });
  }

  if (state.highlightOn && state.compare.size > 0) {
    const highlight = arr
      .map((s, i) => (state.compare.has(s.name) ? { s, x: xVals[i] } : null))
      .filter(Boolean);
    if (highlight.length) {
      traces.push({
        type: "scattergl",
        mode: "markers+text",
        name: "Сравнение",
        x: highlight.map((p) => p.x),
        y: highlight.map((p) => p.s.accepted),
        text: highlight.map((p) => p.s.name),
        textposition: "top center",
        textfont: { size: 12, color: getCss("--ink") },
        marker: {
          size: 12,
          symbol: "diamond",
          color: "rgba(255,255,255,0.1)",
          line: { width: 2, color: getCss("--accent") },
        },
        hovertemplate: "<b>%{text}</b><extra></extra>",
        showlegend: true,
      });
    }
  }

  if (state.selected) {
    const found = arr.findIndex((s) => s.name === state.selected.name);
    if (found >= 0) {
      const sx = xVals[found];
      traces.push({
        type: "scattergl",
        mode: "markers",
        name: "Выбрано",
        x: [sx],
        y: [state.selected.accepted],
        marker: {
          size: 16,
          symbol: "circle-open",
          color: "rgba(255,255,255,0)",
          line: { width: 3, color: getCss("--accent") },
        },
        hoverinfo: "skip",
        showlegend: false,
      });
    }
  }

  return traces;
}

function plotThemeBase() {
  const isDark = document.documentElement.getAttribute("data-theme") === "dark";
  const grid = isDark ? "rgba(255,255,255,0.12)" : "rgba(28,37,49,0.12)";
  const axis = isDark ? "rgba(240,244,248,0.75)" : "rgba(28,37,49,0.75)";
  return { grid, axis };
}

function renderScatter(arr) {
  const traces = buildScatterTraces(arr);
  const { grid, axis } = plotThemeBase();
  const title =
    state.rankMode === "position"
      ? "Позиция в рейтинге (1 — лучший)"
      : state.rankMode === "competition"
      ? "Ранг с учетом равенства (competition)"
      : "Ранг по уровням зачета (dense)";

  const layout = {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    hovermode: "closest",
    height: 560,
    margin: { l: 60, r: 20, t: 40, b: 55 },
    legend: {
      orientation: "h",
      x: 0,
      y: 1.08,
      bgcolor: "rgba(0,0,0,0)",
      font: { size: 11, color: axis },
    },
    xaxis: {
      title: { text: title, font: { size: 12, color: axis } },
      gridcolor: grid,
      zeroline: false,
      tickfont: { color: axis },
      linecolor: grid,
      showline: true,
    },
    yaxis: {
      title: { text: `Зачтено (из ${state.hwCount})`, font: { size: 12, color: axis } },
      gridcolor: grid,
      zeroline: false,
      tickfont: { color: axis },
      linecolor: grid,
      showline: true,
      rangemode: "tozero",
      dtick: 1,
    },
  };

  Plotly.react(el.plotScatter, traces, layout, {
    responsive: true,
    displaylogo: false,
    modeBarButtonsToRemove: ["select2d", "lasso2d", "autoScale2d"],
  });

  if (!state.plotBound) {
    el.plotScatter.on("plotly_click", (ev) => {
      const pt = ev.points?.[0];
      if (!pt) return;
      const name = pt.customdata?.[0];
      if (name) setSelectedByName(name);
    });
    state.plotBound = true;
  }
}

function renderDist(arr) {
  const base = state.distMode === "total" ? state.students : arr;
  const counts = new Array(state.hwCount + 1).fill(0);
  base.forEach((s) => {
    counts[s.accepted] += 1;
  });
  const { grid, axis } = plotThemeBase();
  const x = counts.map((_, i) => i);
  const y = counts;
  const layout = {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    height: 280,
    margin: { l: 50, r: 20, t: 30, b: 45 },
    xaxis: {
      title: { text: "Зачтено работ", font: { size: 11, color: axis } },
      gridcolor: grid,
      tickfont: { color: axis },
      linecolor: grid,
    },
    yaxis: {
      title: { text: "Студентов", font: { size: 11, color: axis } },
      gridcolor: grid,
      tickfont: { color: axis },
      linecolor: grid,
    },
  };
  Plotly.react(
    el.plotDist,
    [
      {
        type: "bar",
        x,
        y,
        marker: { color: getCss("--accent") },
        hovertemplate: "%{x} зачтено: %{y}<extra></extra>",
      },
    ],
    layout,
    { responsive: true, displaylogo: false }
  );
}

function cumulativeFor(student) {
  const out = [];
  let acc = 0;
  for (let i = 0; i < state.hwCount; i++) {
    acc += student.per_hw[i] === 1 ? 1 : 0;
    out.push(acc);
  }
  return out;
}

function averageCumulative(students) {
  const n = students.length || 1;
  const totals = new Array(state.hwCount).fill(0);
  students.forEach((s) => {
    const series = cumulativeFor(s);
    for (let i = 0; i < series.length; i++) {
      totals[i] += series[i];
    }
  });
  return totals.map((v) => Math.round((v / n) * 100) / 100);
}

function renderCompare() {
  const { grid, axis } = plotThemeBase();
  const labels = state.hws.map((hw) => hw.label);
  const avg = averageCumulative(state.students);
  const traces = [
    {
      type: "scatter",
      mode: "lines+markers",
      name: "Среднее по группе",
      x: labels,
      y: avg,
      line: { dash: "dot", color: getCss("--muted") },
      marker: { size: 6 },
      hovertemplate: "%{x}: %{y}<extra></extra>",
    },
  ];

  const colors = [
    "#1f8a70",
    "#f2b134",
    "#e06f59",
    "#3c5a85",
    "#6f7c3a",
    "#1b9aaa",
  ];
  let colorIdx = 0;
  [...state.compare].forEach((name) => {
    const student = state.students.find((s) => s.name === name);
    if (!student) return;
    traces.push({
      type: "scatter",
      mode: "lines+markers",
      name: student.name,
      x: labels,
      y: cumulativeFor(student),
      line: { width: 2.5, color: colors[colorIdx % colors.length] },
      marker: { size: 6 },
      hovertemplate: "%{x}: %{y}<extra></extra>",
    });
    colorIdx += 1;
  });

  const layout = {
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    height: 300,
    margin: { l: 50, r: 20, t: 25, b: 45 },
    legend: {
      orientation: "h",
      x: 0,
      y: 1.05,
      bgcolor: "rgba(0,0,0,0)",
      font: { size: 10, color: axis },
    },
    xaxis: {
      title: { text: "Домашние работы", font: { size: 11, color: axis } },
      gridcolor: grid,
      tickfont: { color: axis },
      linecolor: grid,
    },
    yaxis: {
      title: { text: "Кумулятивный зачет", font: { size: 11, color: axis } },
      gridcolor: grid,
      tickfont: { color: axis },
      linecolor: grid,
      rangemode: "tozero",
      dtick: 1,
    },
  };

  Plotly.react(el.plotCompare, traces, layout, {
    responsive: true,
    displaylogo: false,
  });
}

function statusClass(status) {
  if (status === "хорошие показатели") return "good";
  if (status === "средние показатели") return "mid";
  if (status === "низкие показатели") return "bad";
  return "bad";
}

function renderDetail() {
  if (!state.selected) {
    el.detail.innerHTML = `<div class="detail-empty">Выберите точку или имя, чтобы увидеть персональную карточку.</div>`;
    el.selPill.textContent = "ничего не выбрано";
    el.btnAddCompare.disabled = true;
    return;
  }
  const avgAccepted = state.students.length
    ? state.students.reduce((a, b) => a + b.accepted, 0) / state.students.length
    : 0;
  const delta = state.selected.accepted - avgAccepted;
  const deltaLabel = delta >= 0 ? `+${delta.toFixed(1)}` : delta.toFixed(1);

  const hwCells = state.hws
    .map((hw, idx) => {
      const value = state.selected.per_hw[idx];
      const raw = state.selected.per_hw_raw[idx] || "";
      let stateKey = "na";
      if (value === 1) stateKey = "ok";
      if (value === 0) stateKey = "bad";
      return `<div class="hw-chip" data-state="${stateKey}" title="${escHtml(raw)}">${hw.label}</div>`;
    })
    .join("");

  el.detail.innerHTML = `
    <div class="detail-name">${escHtml(state.selected.name)}</div>
    <div class="detail-meta">
      <span class="badge">позиция: <b>${state.selected.rank}</b></span>
      <span class="badge">зачтено: <b>${state.selected.accepted}/${state.hwCount}</b></span>
      <span class="badge">${state.selected.percent}%</span>
      <span class="badge ${statusClass(state.selected.status)}">статус: <b>${escHtml(
    state.selected.status
  )}</b></span>
      <span class="badge">от ср. группы: <b>${deltaLabel}</b></span>
    </div>
    <div class="hw-grid">${hwCells}</div>
  `;
  el.selPill.textContent = `выбран: ${state.selected.name}`;
  el.btnAddCompare.disabled = false;
}

function renderCompareList() {
  if (state.compare.size === 0) {
    el.compareList.innerHTML = `<div class="detail-empty">Список сравнения пуст. Добавьте людей через поиск.</div>`;
    return;
  }
  const items = [...state.compare]
    .map((name) => state.students.find((s) => s.name === name))
    .filter(Boolean)
    .sort((a, b) => a.rank - b.rank);
  el.compareList.innerHTML = items
    .map(
      (s) => `
    <div class="compare-item">
      <div class="compare-item-title">${escHtml(s.name)}</div>
      <div class="compare-item-meta">
        <span>позиция: ${s.rank}</span>
        <span>зачтено: ${s.accepted}/${state.hwCount}</span>
        <span class="${statusClass(s.status)}">${escHtml(s.status)}</span>
        <button class="compare-remove" data-name="${escHtml(s.name)}">убрать</button>
      </div>
    </div>
  `
    )
    .join("");

  el.compareList.querySelectorAll(".compare-remove").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.getAttribute("data-name");
      if (name) {
        state.compare.delete(name);
        saveCompare();
        render();
      }
    });
  });
}

function render() {
  const arr = applyFilters();
  const stats = computeStats(arr);
  el.kTotal.textContent = String(state.students.length);
  el.kFiltered.textContent = String(stats.n);
  el.kMean.textContent = stats.n ? stats.mean.toFixed(2) : "—";
  el.kGoodShare.textContent = stats.n ? `${(stats.goodShare * 100).toFixed(1)}%` : "—";
  el.kFilteredSub.textContent = `мин=${state.minAccepted}, статусы=${[...state.statuses].join(
    ", "
  )}`;

  renderScatter(arr);
  renderDist(arr);
  renderCompare();
  renderDetail();
  renderCompareList();
}

function setSelectedByName(name) {
  const rec = state.students.find((s) => s.name === name);
  if (!rec) return;
  state.selected = rec;
  render();
}

function hideSuggest(box) {
  box.style.display = "none";
  box.innerHTML = "";
}

function showSuggest(box, items, onPick) {
  if (!items.length) {
    hideSuggest(box);
    return;
  }
  box.innerHTML = items
    .map((nm) => `<button type="button" data-name="${escHtml(nm)}">${escHtml(nm)}</button>`)
    .join("");
  box.style.display = "block";
  box.querySelectorAll("button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.getAttribute("data-name");
      if (name) onPick(name);
    });
  });
}

const doSuggest = debounce(() => {
  const q = (el.q.value || "").trim().toLowerCase();
  if (!q) {
    hideSuggest(el.suggest);
    return;
  }
  const matches = state.students
    .filter((s) => s.name.toLowerCase().includes(q))
    .slice(0, 12)
    .map((s) => s.name);
  showSuggest(el.suggest, matches, (name) => {
    setSelectedByName(name);
    el.q.value = name;
    hideSuggest(el.suggest);
  });
}, 120);

const doCompareSuggest = debounce(() => {
  const q = (el.compareInput.value || "").trim().toLowerCase();
  if (!q) {
    hideSuggest(el.compareSuggest);
    return;
  }
  const matches = state.students
    .filter((s) => s.name.toLowerCase().includes(q))
    .slice(0, 12)
    .map((s) => s.name);
  showSuggest(el.compareSuggest, matches, (name) => {
    state.compare.add(name);
    saveCompare();
    el.compareInput.value = "";
    hideSuggest(el.compareSuggest);
    render();
  });
}, 120);

document.querySelectorAll("#chipsStatus .chip").forEach((chip) => {
  chip.addEventListener("click", () => {
    const key = chip.getAttribute("data-key");
    if (!key) return;
    const active = chip.getAttribute("data-active") === "true";
    chip.setAttribute("data-active", String(!active));
    if (active) state.statuses.delete(key);
    else state.statuses.add(key);
    render();
  });
});

el.minAcc.addEventListener("input", () => {
  state.minAccepted = Number(el.minAcc.value);
  el.minAccLabel.textContent = String(state.minAccepted);
  render();
});

el.topN.addEventListener("change", () => {
  state.topN = Number(el.topN.value);
  render();
});

el.rankMode.addEventListener("change", () => {
  state.rankMode = el.rankMode.value;
  render();
});

el.colorMode.addEventListener("change", () => {
  state.colorMode = el.colorMode.value;
  render();
});

el.distMode.addEventListener("change", () => {
  state.distMode = el.distMode.value;
  render();
});

el.btnTheme.addEventListener("click", () => {
  const next =
    document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
  setTheme(next);
});

el.btnReset.addEventListener("click", () => {
  state.statuses = new Set(["хорошие показатели", "средние показатели", "низкие показатели"]);
  document.querySelectorAll("#chipsStatus .chip").forEach((chip) => {
    chip.setAttribute("data-active", "true");
  });
  state.minAccepted = 0;
  el.minAcc.value = "0";
  el.minAccLabel.textContent = "0";
  state.topN = 0;
  el.topN.value = "0";
  state.rankMode = "position";
  el.rankMode.value = "position";
  state.colorMode = "status";
  el.colorMode.value = "status";
  state.distMode = "filtered";
  el.distMode.value = "filtered";
  el.q.value = "";
  hideSuggest(el.suggest);
  render();
});

el.btnExport.addEventListener("click", () => {
  const arr = applyFilters();
  const headers = ["rank", "accepted", "percent", "status", "name"];
  const lines = [headers.join(",")];
  arr.forEach((s) => {
    const row = headers.map((h) => {
      const value = s[h];
      const text = String(value ?? "");
      if (text.includes('"') || text.includes(",") || text.includes("\n")) {
        return `"${text.replaceAll('"', '""')}"`;
      }
      return text;
    });
    lines.push(row.join(","));
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "rating_filtered.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

el.btnRefresh.addEventListener("click", async () => {
  el.btnRefresh.disabled = true;
  el.btnRefresh.textContent = "Обновление...";
  try {
    await fetchData(true);
  } catch (err) {
    console.error(err);
  } finally {
    el.btnRefresh.disabled = false;
    el.btnRefresh.textContent = "Обновить XLSX";
  }
});

el.btnHighlight.addEventListener("click", () => {
  state.highlightOn = !state.highlightOn;
  el.btnHighlight.textContent = `Подсветка: ${state.highlightOn ? "включена" : "выключена"}`;
  render();
});

el.btnAddCompare.addEventListener("click", () => {
  if (!state.selected) return;
  state.compare.add(state.selected.name);
  saveCompare();
  render();
});

el.btnClearCompare.addEventListener("click", () => {
  state.compare.clear();
  saveCompare();
  render();
});

el.q.addEventListener("input", doSuggest);
el.q.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    const q = (el.q.value || "").trim().toLowerCase();
    if (!q) return;
    const best = state.students.find((s) => s.name.toLowerCase().includes(q));
    if (best) {
      setSelectedByName(best.name);
      hideSuggest(el.suggest);
    }
  }
  if (e.key === "Escape") {
    hideSuggest(el.suggest);
  }
});

el.compareInput.addEventListener("input", doCompareSuggest);
el.compareInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    const q = (el.compareInput.value || "").trim().toLowerCase();
    if (!q) return;
    const best = state.students.find((s) => s.name.toLowerCase().includes(q));
    if (best) {
      state.compare.add(best.name);
      saveCompare();
      el.compareInput.value = "";
      hideSuggest(el.compareSuggest);
      render();
    }
  }
  if (e.key === "Escape") {
    hideSuggest(el.compareSuggest);
  }
});

document.addEventListener("click", (e) => {
  if (!el.suggest.contains(e.target) && e.target !== el.q) {
    hideSuggest(el.suggest);
  }
  if (!el.compareSuggest.contains(e.target) && e.target !== el.compareInput) {
    hideSuggest(el.compareSuggest);
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key === "/" && document.activeElement !== el.q) {
    e.preventDefault();
    el.q.focus();
  }
});

loadCompare();
loadTheme();
fetchData().catch((err) => {
  console.error(err);
});
