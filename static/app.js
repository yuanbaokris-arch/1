"use strict";

(() => {
  const $ = (id) => document.getElementById(id);
  const app = {
    state: { dataset: null, version: null, draft: null },
    selected: new Set(), locked: new Set(), busy: false, warningsOpen: false,
    month: "1月", department: "", search: "", changedOnly: false, goal: null,
  };
  const numberFormat = new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 3 });
  const escape = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const num = (value) => value === null || value === undefined || value === "" || !Number.isFinite(Number(value)) ? null : Number(value);
  const fmt = (value) => num(value) === null ? "—" : numberFormat.format(Number(value));
  const signed = (value) => num(value) === null ? "—" : value > 0 ? `+${fmt(value)}` : fmt(value);
  const equal = (a, b) => num(a) !== null && num(b) !== null ? Math.abs(Number(a) - Number(b)) < 0.000001 : a === b;
  const activeDraft = () => app.state.draft && app.state.draft.parent_id === app.state.version?.id ? app.state.draft : null;
  const current = () => activeDraft() || app.state.version;
  const allRows = () => Array.isArray(current()?.rows) ? current().rows : [];
  const operations = () => Array.isArray(current()?.operations) ? current().operations : [];
  const changed = (row) => !equal(row.vehicles, row.original_vehicles) || !equal(row.monthly_trips, row.original_monthly_trips);
  const visibleRows = () => allRows().filter((row) => (!app.month || String(row.month) === app.month) && (!app.department || row.department === app.department) && (!app.search || `${row.route} ${row.department} ${row.fleet}`.toLocaleLowerCase().includes(app.search.toLocaleLowerCase())) && (!app.changedOnly || changed(row)));
  const sum = (rows, key) => rows.reduce((total, row) => total + (num(row[key]) ?? 0), 0);
  const dateText = (value, timeOnly = false) => {
    if (!value) return "时间未记录";
    const date = new Date(value);
    if (!Number.isFinite(date.getTime())) return String(value);
    return new Intl.DateTimeFormat("zh-CN", { timeZone: "Asia/Shanghai", ...(timeOnly ? {} : { month: "2-digit", day: "2-digit" }), hour: "2-digit", minute: "2-digit", hour12: false }).format(date);
  };
  const lockIcon = (locked) => `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.4" aria-hidden="true"><rect x="5" y="9" width="10" height="8" rx="1.5"/><path d="${locked ? "M7 9V6a3 3 0 0 1 6 0v3" : "M7 9V6a3 3 0 0 1 6-1"}"/><path d="M10 12v2"/></svg>`;

  function toast(message, error = false) {
    const item = document.createElement("div");
    item.className = `toast${error ? " error" : ""}`;
    item.textContent = message;
    item.setAttribute("role", error ? "alert" : "status");
    $("toast-region").append(item);
    while ($("toast-region").children.length > 3) $("toast-region").firstElementChild.remove();
    window.setTimeout(() => item.remove(), error ? 8500 : 4500);
  }

  async function request(path, options = {}) {
    const init = { ...options };
    if (init.body && !(init.body instanceof FormData)) {
      init.headers = { "Content-Type": "application/json", ...init.headers };
      init.body = JSON.stringify(init.body);
    }
    const response = await fetch(path, init);
    const body = await response.json().catch(() => null);
    if (!response.ok) {
      const detail = body?.detail;
      const error = new Error(typeof detail === "string" ? detail : detail?.message || `操作未完成（${response.status}），请稍后重试。`);
      error.status = response.status;
      throw error;
    }
    return body;
  }

  async function run(title, task, message = "正在使用原表公式重新计算，可能需要一些时间，请勿重复操作。") {
    if (app.busy) return;
    app.busy = true;
    $("busy-title").textContent = title;
    $("busy-message").textContent = message;
    $("busy-overlay").hidden = false;
    document.body.setAttribute("aria-busy", "true");
    document.querySelectorAll("dialog[open]").forEach((dialog) => dialog.close());
    try { await task(); }
    catch (error) { toast(error.message || "连接未成功，请确认服务仍在运行后重试。", true); }
    finally {
      app.busy = false;
      $("busy-overlay").hidden = true;
      document.body.removeAttribute("aria-busy");
      renderSelection();
    }
  }

  function readPreferences() {
    try {
      const saved = JSON.parse(localStorage.getItem(`bus-budget:${app.state.dataset.id}`) || "{}");
      app.locked = new Set(Array.isArray(saved.locked) ? saved.locked.map(String) : []);
      app.goal = num(saved.goal);
      $("trip-target").value = app.goal ?? "";
    } catch { app.locked = new Set(); app.goal = null; }
  }

  function savePreferences() {
    if (!app.state.dataset) return;
    try { localStorage.setItem(`bus-budget:${app.state.dataset.id}`, JSON.stringify({ locked: [...app.locked], goal: app.goal })); }
    catch { /* A blocked browser storage never prevents a calculation. */ }
  }

  function applyState(state) {
    const newDataset = state.dataset?.id !== app.state.dataset?.id;
    app.state = state;
    if (newDataset) {
      app.selected.clear(); app.month = "1月"; app.department = ""; app.search = ""; app.changedOnly = false;
      $("route-search").value = ""; $("changed-only").checked = false;
      if (state.dataset) readPreferences();
    }
    const keys = new Set(allRows().map((row) => row.key));
    app.selected = new Set([...app.selected].filter((key) => keys.has(key)));
    render();
  }

  function render() {
    const dataset = app.state.dataset;
    $("demo-button").textContent = app.state.capabilities?.demo_available ? "打开已提供的预算基础表 →" : "选择预算基础表 →";
    $("upload-zone").querySelector("strong").textContent = app.state.capabilities?.demo_available ? "或导入你的预算基础表" : "导入你的预算基础表";
    $("empty-state").hidden = !!dataset;
    $("workspace").hidden = !dataset;
    $("heading-actions").hidden = !dataset;
    $("versions-button").hidden = !dataset;
    if (!dataset) return;
    $("dataset-name").textContent = dataset.name || "预算基础表.xlsx";
    $("dataset-description").textContent = `${dataset.route_count || new Set(allRows().map((r) => r.route)).size} 条线路 · ${(dataset.months || []).length || 12} 个月份 · 原始文件已独立保留`;
    $("baseline-status").textContent = dataset.baseline_verified === true ? "空操作重算对照一致" : dataset.baseline_verified === false ? "重算对照存在差异" : "保留原表公式";
    $("baseline-status").className = `tag ${dataset.baseline_verified === false ? "tag-neutral" : "tag-green"}`;
    $("version-count").textContent = (dataset.versions || []).length;
    const months = dataset.months?.length ? dataset.months : [...new Set(allRows().map((r) => r.month))];
    if (app.month && !months.includes(app.month)) app.month = months[0] || "";
    $("month-filter").innerHTML = months.map((month) => `<option value="${escape(month)}">${escape(String(month).replace("月", " 月"))}</option>`).join("") + '<option value="">全年 · 全部月份</option>';
    $("month-filter").value = app.month;
    const departments = [...new Set(allRows().map((r) => r.department).filter(Boolean))];
    $("department-filter").innerHTML = '<option value="">全部业务部</option>' + departments.map((dep) => `<option value="${escape(dep)}">${escape(dep)}</option>`).join("");
    $("department-filter").value = app.department;
    for (let i = 1; i <= 4; i++) {
      $("step-" + i).classList.toggle("complete", i === 1 || (i === 2 && operations().length > 0));
      $("step-" + i).classList.toggle("active", i === (activeDraft() ? 3 : (app.state.version?.label !== "V0" && app.state.version?.label ? 4 : 2)));
    }
    renderTable(); renderMetrics(); renderDraft(); renderWarnings(); renderChart(); renderVersions();
  }

  function renderTable() {
    const rows = visibleRows();
    $("routes-body").innerHTML = rows.map((row) => {
      const locked = app.locked.has(String(row.route));
      const delta = num(row.monthly_trips) !== null && num(row.original_monthly_trips) !== null ? row.monthly_trips - row.original_monthly_trips : null;
      const value = (key, originalKey) => `<span class="${num(row[key]) < 0 ? "bad-number " : ""}${originalKey && !equal(row[key], row[originalKey]) ? "changed-value " : ""}value-current">${originalKey && !equal(row[key], row[originalKey]) ? `<span class="value-previous">${fmt(row[originalKey])}</span>` : ""}${fmt(row[key])}</span>`;
      return `<tr class="${app.selected.has(row.key) ? "selected " : ""}${locked ? "locked" : ""}"><td class="check-cell"><input type="checkbox" data-select="${escape(row.key)}" aria-label="选择 ${escape(row.month)} ${escape(row.route)}" ${app.selected.has(row.key) ? "checked" : ""} ${locked ? "disabled" : ""}></td><td><span class="route-name" title="${escape(row.route)}">${escape(row.route)}</span><span class="route-detail">${escape(row.department || "未分组")}${row.fleet ? ` · ${escape(row.fleet)}` : ""}</span></td><td class="month-cell">${escape(row.month)}</td><td class="numeric" title="${escape(row.month)}!J${escape(row.source_row)} · 仅此投放输入参与减车">${value("vehicles", "original_vehicles")}</td><td class="numeric">${value("day_trips", "original_day_trips")}</td><td class="numeric">${value("a_days")}</td><td class="numeric">${value("monthly_trips")}</td><td class="numeric"><span class="delta ${delta < 0 ? "delta-negative" : delta > 0 ? "delta-positive" : "delta-zero"}">${delta === 0 ? "—" : signed(delta)}</span></td><td class="lock-cell"><button class="lock-button ${locked ? "is-locked" : ""}" data-lock="${escape(row.route)}" aria-label="${locked ? "解锁" : "锁定"}${escape(row.route)}所有月份" aria-pressed="${locked}" title="${locked ? "已锁定，点击解锁" : "锁定该线路所有月份"}">${lockIcon(locked)}</button></td></tr>`;
    }).join("");
    $("table-empty").hidden = rows.length > 0;
    const count = new Set(rows.map((row) => row.route)).size;
    $("route-count").textContent = `${fmt(count)} 条线路${app.month ? "" : ` · ${fmt(rows.length)} 条月度记录`}${app.locked.size ? ` · ${app.locked.size} 条线路已锁定` : ""}`;
    renderSelection();
  }

  function renderSelection() {
    const rows = visibleRows();
    const eligible = rows.filter((row) => !app.locked.has(String(row.route)));
    const selected = allRows().filter((row) => app.selected.has(row.key) && !app.locked.has(String(row.route)));
    const selectedVisible = eligible.filter((row) => app.selected.has(row.key));
    $("select-all").checked = eligible.length > 0 && selectedVisible.length === eligible.length;
    $("select-all").indeterminate = selectedVisible.length > 0 && selectedVisible.length < eligible.length;
    $("select-all").disabled = eligible.length === 0 || app.busy;
    $("selection-count").textContent = selected.length ? `已选择 ${new Set(selected.map((row) => row.route)).size} 条线路 · ${selected.length} 条月度记录` : "未选择线路";
    const hiddenCount = selected.filter((row) => !rows.some((r) => r.key === row.key)).length;
    $("selection-note").textContent = hiddenCount ? `其中 ${hiddenCount} 条不在当前筛选中，操作会包含所有已选记录。` : selected.length ? "只调整 A 计划营运1投放；可局部恢复或撤销。" : "勾选要调整的线路，可先锁定重点线路。";
    if (app.state.dataset?.baseline_verified === false) $("selection-note").textContent = "空操作重算对照尚未通过，请查看数据检查，暂不能调整。";
    $("clear-selection").hidden = app.selected.size === 0;
    $("adjust-button").disabled = !selected.length || app.busy || app.state.dataset?.baseline_verified === false;
    $("restore-button").disabled = !selected.length || !activeDraft() || app.busy;
    $("undo-button").disabled = !activeDraft()?.operations?.length || app.busy;
    $("save-button").disabled = !activeDraft()?.operations?.length || app.busy;
  }

  function renderMetrics() {
    const rows = visibleRows();
    const original = sum(rows, "original_monthly_trips"), currentTotal = sum(rows, "monthly_trips"), reduction = original - currentTotal;
    $("metric-original").textContent = fmt(original);
    $("metric-current").textContent = fmt(currentTotal);
    $("metric-current-note").textContent = reduction > 0 ? `较原始减少 ${fmt(reduction)} 趟` : reduction < 0 ? `较原始增加 ${fmt(-reduction)} 趟` : "与原始版本一致";
    $("metric-reduction").textContent = fmt(reduction);
    const count = rows.filter(changed).length;
    $("metric-affected").textContent = count ? `${new Set(rows.filter(changed).map((r) => r.route)).size} 条线路 · ${count} 条月度记录已变化` : "当前范围尚未调整";
    if (app.goal !== null && app.goal > 0) {
      const remaining = app.goal - reduction;
      $("target-result").textContent = remaining > 0 ? `当前范围距目标还需减少 ${fmt(remaining)} 趟` : remaining < 0 ? `当前范围已超过目标 ${fmt(-remaining)} 趟` : "当前筛选范围恰好达到目标";
    } else $("target-result").textContent = "填写后可查看当前筛选范围的差额";
  }

  function operationDetails(operation) {
    const keys = operation.keys || operation.scope?.map((change) => change.key) || operation.changes?.map((change) => change.key) || [];
    const rows = allRows().filter((row) => keys.includes(row.key));
    const routes = operation.routes || [...new Set(rows.map((r) => r.route))];
    const months = operation.months || [...new Set(rows.map((r) => r.month))];
    const type = String(operation.type || operation.kind || operation.action || "");
    const restore = /restore|恢复/.test(type);
    return {
      title: restore ? "局部恢复投放" : "A计划营运1减少1台",
      scope: routes.length ? `${routes.slice(0, 5).join("、")}${routes.length > 5 ? ` 等${routes.length}条线路` : ""}${months.length ? ` · ${months.join("、")}` : ""}${num(operation.trip_delta) !== null ? ` · 趟次${signed(operation.trip_delta)}` : ""}` : `${keys.length || operation.affected_rows || operation.count || "已选"} 条月度记录`,
      time: operation.created_at || operation.timestamp || operation.at,
      reason: operation.reason,
    };
  }

  function renderDraft() {
    const draft = activeDraft();
    const version = app.state.version;
    const label = version?.label || "V0";
    $("current-version-label").textContent = draft ? `${label} 的调整草稿 · 未保存` : `${label}${label === "V0" ? " · 原始版本" : " · 已保存"}`;
    $("draft-title").textContent = draft ? "当前调整草稿" : label === "V0" ? "原始版本 V0" : `方案版本 ${label}`;
    $("draft-status").textContent = draft ? "未保存" : label === "V0" ? "只读" : "已保存";
    $("draft-status").classList.toggle("unsaved", !!draft);
    const changedRows = allRows().filter(changed);
    $("draft-summary").textContent = draft ? `已对 ${new Set(changedRows.map((row) => row.route)).size} 条线路进行试算。核对结果后，填写原因并保存为独立版本。` : label === "V0" ? "选择线路和月份后开始调整，原始数据会完整保留。" : version.reason || "该方案已保存，可从此版本继续创建新方案。";
    $("draft-parent").textContent = draft ? `基于 ${label} · ${operations().length} 次操作` : version?.created_at ? `${dateText(version.created_at)} · 北京时间` : "";
    const ops = operations();
    $("operation-list").innerHTML = ops.length ? [...ops].reverse().map((operation) => {
      const detail = operationDetails(operation);
      return `<article class="operation-item"><div class="operation-title"><strong>${escape(detail.title)}</strong>${detail.time ? `<time>${escape(dateText(detail.time, true))}</time>` : ""}</div><p>${escape(detail.scope)}</p>${detail.reason ? `<small>${escape(detail.reason)}</small>` : ""}</article>`;
    }).join("") : '<div class="operation-empty"><span aria-hidden="true">↗</span><p>还没有调整记录</p><small>每次操作都会记录在这里</small></div>';
    $("save-button").textContent = "保存为新版本 →";
  }

  function warningItems() {
    const source = current()?.warnings ?? app.state.dataset?.warnings ?? [];
    return Array.isArray(source) ? source : [];
  }

  function renderWarnings() {
    const warnings = warningItems();
    const existing = warnings.filter((warning) => warning.origin !== "new");
    const added = warnings.filter((warning) => warning.origin === "new");
    const blocking = added.filter((warning) => warning.severity === "error");
    $("validation-summary").innerHTML = `<div class="validation-pill${existing.length ? "" : " clear"}"><span class="pill-dot"></span>原表已有 <strong>${fmt(existing.length)}</strong> 项提示</div><div class="validation-pill${added.length ? " new" : " clear"}"><span class="pill-dot"></span>本次新增 <strong>${fmt(added.length)}</strong> 项${blocking.length ? `（${blocking.length} 项错误）` : "提示"}</div>`;
    $("warnings-list").hidden = !app.warningsOpen;
    $("toggle-warnings").textContent = app.warningsOpen ? "收起明细 ↑" : "展开明细 ↓";
    $("warnings-list").innerHTML = warnings.length ? [...added, ...existing].map((warning) => `<div class="warning-item"><span class="warning-origin${warning.origin === "new" ? " new" : ""}">${warning.origin === "new" ? "新增" : "原有"}</span><div><p>${escape(warning.message || warning.code || "待核查项目")}</p>${warning.sheet || warning.cell ? `<span class="warning-cell">${escape(warning.sheet || "")}${warning.cell ? ` ! ${escape(warning.cell)}` : ""}</span>` : ""}</div></div>`).join("") : '<div class="warning-item"><p>当前检查未发现提示。最低服务标准等未确认规则尚未参与校验。</p></div>';
  }

  function renderChart() {
    const rows = allRows();
    const months = app.state.dataset.months || [...new Set(rows.map((row) => row.month))];
    const values = months.map((month) => {
      const entries = rows.filter((row) => row.month === month);
      return sum(entries, "original_monthly_trips") - sum(entries, "monthly_trips");
    });
    const max = Math.max(...values.map(Math.abs), 1);
    const width = 270, chartTop = 16, baseline = 82, step = 21;
    $("monthly-chart").innerHTML = `<svg viewBox="0 0 ${width} 116" role="img" aria-label="各月相较原始版本减少的趟次：${escape(months.map((m, i) => `${m}${fmt(values[i])}趟`).join("，"))}"><line x1="7" x2="264" y1="${baseline}" y2="${baseline}" stroke="#e6ecdf"/><line x1="7" x2="264" y1="${chartTop}" y2="${chartTop}" stroke="#eff2e9" stroke-dasharray="2 3"/>${values.map((value, index) => {
      const height = Math.abs(value) / max * 53;
      const x = 12 + index * step;
      const barHeight = Math.max(height, 2);
      return `<rect x="${x}" y="${baseline - barHeight}" width="11" height="${barHeight}" rx="2" fill="${value < 0 ? "#c3ad77" : value === 0 ? "#dfe8d6" : "#84aa7b"}"><title>${escape(months[index])}：${value >= 0 ? "减少" : "增加"}${fmt(Math.abs(value))}趟</title></rect><text x="${x + 5.5}" y="99" text-anchor="middle" fill="#a8b49a" font-size="8">${escape(String(months[index]).replace("月", ""))}</text>`;
    }).join("")}<text x="260" y="9" text-anchor="end" fill="#a8b49a" font-size="7">${values.some((v) => v !== 0) ? `最大 ${fmt(max)}` : "尚无变化"}</text></svg>`;
  }

  function renderVersions() {
    const versions = app.state.dataset?.versions || [];
    $("versions-list").innerHTML = [...versions].reverse().map((version) => {
      const selected = version.id === app.state.version?.id;
      return `<article class="version-item"><span class="version-icon">${escape(version.label || "V?")}</span><div class="version-info"><strong>${escape(version.label === "V0" ? "原始版本 · 不可覆盖" : `${version.label || "方案版本"}${version.parent_id ? " · 已保存方案" : ""}`)}</strong><p>${escape(version.reason || (version.label === "V0" ? "导入时保存的独立原始版本" : "未填写说明"))}</p><small>${escape(dateText(version.created_at))} · 北京时间${num(version.totals?.trip_delta) !== null ? ` · 月趟次变化 ${signed(version.totals.trip_delta)}` : ""}</small></div><button class="button button-outline" data-version="${escape(version.id)}" ${selected ? "disabled" : ""}>${selected ? "当前版本" : "查看"}</button></article>`;
    }).join("") || '<div class="table-empty">还没有已保存版本。</div>';
  }

  async function importFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) return toast("请选择 .xlsx 格式的预算基础表。", true);
    if (app.state.dataset && !window.confirm("导入另一份基础表将切换到新的数据集。请确认当前未保存草稿已保存。继续导入？")) return;
    const form = new FormData(); form.append("file", file);
    await run("正在导入预算基础表", async () => {
      applyState(await request("/api/import", { method: "POST", body: form }));
      toast("导入完成，原始版本 V0 已独立保留。请先查看数据检查。 ");
    }, "正在识别线路、月份和原表公式，并检查已有数据问题。");
    $("file-input").value = "";
  }

  async function ensureDraft() {
    if (activeDraft()) return true;
    const existing = app.state.draft;
    let discard = false;
    if (existing?.operations?.length) {
      discard = window.confirm("另一个版本仍有未保存草稿。继续会放弃那份草稿及其中的调整记录，已保存版本不受影响。确定放弃并从当前版本开始？");
      if (!discard) return false;
    }
    const selected = new Set(app.selected);
    applyState(await request("/api/draft", { method: "POST", body: { version_id: app.state.version.id, ...(discard ? { discard: true } : {}) } }));
    app.selected = selected;
    return true;
  }

  async function adjust() {
    const keys = [...app.selected].filter((key) => allRows().some((row) => row.key === key && !app.locked.has(String(row.route))));
    if (!keys.length) return;
    await run("正在试算减车影响", async () => {
      if (!await ensureDraft()) return;
      applyState(await request("/api/draft/adjust", { method: "POST", body: { keys, delta: -1, locked_routes: [...app.locked] } }));
      toast(`已对 ${keys.length} 条月度记录试算。请核对趟次变化和新增提示。`);
    });
  }

  function bindEvents() {
    $("help-button").addEventListener("click", () => $("help-dialog").showModal());
    $("versions-button").addEventListener("click", () => { renderVersions(); $("versions-dialog").showModal(); });
    document.querySelectorAll("[data-close-dialog]").forEach((button) => button.addEventListener("click", () => button.closest("dialog").close()));
    document.querySelectorAll("dialog").forEach((dialog) => dialog.addEventListener("click", (event) => { if (event.target === dialog) { const rect = dialog.getBoundingClientRect(); if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) dialog.close(); } }));
    $("demo-button").addEventListener("click", () => {
      if (!app.state.capabilities?.demo_available) { $("file-input").click(); return; }
      run("正在打开已提供的基础表", async () => {
      applyState(await request("/api/demo", { method: "POST", body: {} }));
      toast("已打开独立副本。原表日期及已有异常可在数据检查中查看。");
      }, "正在读取 2027 年预算基础表，保留原始公式和投放数据。");
    });
    $("import-another").addEventListener("click", () => $("file-input").click());
    $("upload-zone").addEventListener("click", () => $("file-input").click());
    $("upload-zone").addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); $("file-input").click(); } });
    $("file-input").addEventListener("change", (event) => importFile(event.target.files[0]));
    for (const type of ["dragenter", "dragover"]) $("upload-zone").addEventListener(type, (event) => { event.preventDefault(); $("upload-zone").classList.add("drag-over"); });
    for (const type of ["dragleave", "drop"]) $("upload-zone").addEventListener(type, (event) => { event.preventDefault(); $("upload-zone").classList.remove("drag-over"); if (type === "drop") importFile(event.dataTransfer.files[0]); });
    const filter = () => { app.selected.clear(); renderTable(); renderMetrics(); };
    $("route-search").addEventListener("input", (event) => { app.search = event.target.value.trim(); filter(); });
    $("month-filter").addEventListener("change", (event) => { app.month = event.target.value; filter(); });
    $("department-filter").addEventListener("change", (event) => { app.department = event.target.value; filter(); });
    $("changed-only").addEventListener("change", (event) => { app.changedOnly = event.target.checked; filter(); });
    $("trip-target").addEventListener("input", (event) => { app.goal = num(event.target.value); if (app.goal !== null && app.goal < 0) { event.target.setCustomValidity("请填写正数目标。"); } else event.target.setCustomValidity(""); savePreferences(); renderMetrics(); });
    $("routes-body").addEventListener("change", (event) => {
      const input = event.target.closest("[data-select]");
      if (!input) return;
      input.checked ? app.selected.add(input.dataset.select) : app.selected.delete(input.dataset.select);
      input.closest("tr").classList.toggle("selected", input.checked);
      renderSelection();
    });
    $("routes-body").addEventListener("click", (event) => {
      const button = event.target.closest("[data-lock]");
      if (!button) return;
      const route = button.dataset.lock;
      if (app.locked.has(route)) app.locked.delete(route);
      else { app.locked.add(route); allRows().filter((row) => String(row.route) === route).forEach((row) => app.selected.delete(row.key)); }
      savePreferences(); renderTable();
      toast(`${route}已${app.locked.has(route) ? "锁定所有月份" : "解锁"}。`);
    });
    $("select-all").addEventListener("change", (event) => { visibleRows().filter((row) => !app.locked.has(String(row.route))).forEach((row) => event.target.checked ? app.selected.add(row.key) : app.selected.delete(row.key)); renderTable(); });
    $("clear-selection").addEventListener("click", () => { app.selected.clear(); renderTable(); });
    $("adjust-button").addEventListener("click", adjust);
    $("restore-button").addEventListener("click", () => run("正在恢复选中线路", async () => {
      const keys = [...app.selected].filter((key) => allRows().some((row) => row.key === key && !app.locked.has(String(row.route))));
      if (!keys.length || !activeDraft()) return;
      applyState(await request("/api/draft/restore", { method: "POST", body: { keys, locked_routes: [...app.locked] } }));
      toast("选中记录已恢复到草稿父版本，原有公式已重新计算。");
    }));
    $("undo-button").addEventListener("click", () => run("正在撤销上一步", async () => { applyState(await request("/api/draft/undo", { method: "POST", body: {} })); toast("已撤销上一步，结果已重新计算。"); }));
    $("save-button").addEventListener("click", () => {
      const changedRows = allRows().filter(changed);
      $("save-description").textContent = `本草稿基于 ${app.state.version.label}，包含 ${operations().length} 次操作，较原始版本有 ${changedRows.length} 条月度记录发生变化。保存后可以继续创建下一版。`;
      $("save-reason").value = "";
      $("save-dialog").showModal();
      $("save-reason").focus();
    });
    $("save-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const reason = $("save-reason").value.trim();
      if (!reason) { $("save-reason").setCustomValidity("请填写调整原因。"); $("save-reason").reportValidity(); return; }
      run("正在保存方案版本", async () => { applyState(await request("/api/draft/save", { method: "POST", body: { reason } })); toast(`${app.state.version?.label || "新版本"} 已保存，修改记录可在版本记录中查看。`); }, "正在保存投放输入、原表公式、调整原因和操作记录。");
    });
    $("save-reason").addEventListener("input", () => $("save-reason").setCustomValidity(""));
    $("toggle-warnings").addEventListener("click", () => { app.warningsOpen = !app.warningsOpen; renderWarnings(); });
    $("versions-list").addEventListener("click", (event) => {
      const button = event.target.closest("[data-version]");
      if (!button) return;
      run("正在读取历史版本", async () => { app.selected.clear(); applyState(await request(`/api/state?version_id=${encodeURIComponent(button.dataset.version)}`)); toast(`已打开 ${app.state.version.label}。${activeDraft() ? "该版本已有未保存草稿。" : "可以查看结果或从此版本继续调整。"}`); }, "正在读取该版本结果，其他已保存版本会完整保留。");
    });
    $("export-button").addEventListener("click", () => {
      if (activeDraft()?.operations?.length) { toast("当前仍有未保存调整，请先保存版本再导出。", true); $("save-button").focus(); return; }
      $("export-description").textContent = `当前导出版本：${app.state.version.label}。结果注明方案版本，原始上传文件不变。`;
      $("export-dialog").showModal();
    });
    document.querySelectorAll("[data-export-kind]").forEach((button) => button.addEventListener("click", () => run("正在生成 Excel", async () => {
      const response = await fetch(`/api/export?version_id=${encodeURIComponent(app.state.version.id)}&kind=${button.dataset.exportKind}`);
      if (!response.ok) { const body = await response.json().catch(() => ({})); throw new Error(typeof body.detail === "string" ? body.detail : "导出失败，请稍后重试。"); }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const utfName = disposition.match(/filename\*=UTF-8''([^;]+)/i);
      const simpleName = disposition.match(/filename="?([^";]+)"?/i);
      let filename = `${app.state.version.label}_${button.dataset.exportKind === "monthly" ? "月趟次试算表" : "预算基础表"}.xlsx`;
      try { if (utfName) filename = decodeURIComponent(utfName[1]); else if (simpleName) filename = simpleName[1]; } catch { /* Keep safe fallback filename. */ }
      const url = URL.createObjectURL(blob), link = document.createElement("a");
      link.href = url; link.download = filename; document.body.append(link); link.click(); link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 10000);
      toast("Excel 已生成并开始下载。 ");
    }, "正在生成含方案版本和原表公式的 Excel 文件。")));
    window.addEventListener("beforeunload", (event) => { if (activeDraft()?.operations?.length) { event.preventDefault(); event.returnValue = ""; } });
  }

  bindEvents();
  run("正在连接预算工作台", async () => { applyState(await request("/api/state")); }, "正在读取当前数据与已保存版本。");
})();
