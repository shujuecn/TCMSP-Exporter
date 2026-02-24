// ==UserScript==
// @name         TCMSP Exporter
// @namespace    https://github.com/shujuecn/TCMSP-Exporter
// @version      1.3.0
// @description  TCMSP-Exporter 批量抓取成分/靶点/疾病并导出 XLSX
// @author       shujuecn + codex
// @match        https://www.tcmsp-e.com/*
// @match        https://tcmsp-e.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @license      MIT
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const SEARCH_URL = "https://www.tcmsp-e.com/tcmspsearch.php";
  const panelId = "tcmsp-spider-panel";
  const QUERY_CONCURRENCY = 3;
  const POS_STORE_KEY = "tcmsp_spider_panel_pos_v1";
  const COLLAPSE_STORE_KEY = "tcmsp_spider_panel_collapsed_v1";
  let cachedToken = "";

  function safeJsonParse(text, fallback) {
    try {
      return JSON.parse(text);
    } catch (_) {
      return fallback;
    }
  }

  function safeEvalLiteral(text, fallback) {
    try {
      return Function(`"use strict"; return (${text});`)();
    } catch (_) {
      return fallback;
    }
  }

  function parseTokenFromHtml(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return doc.querySelector('input[name="token"]')?.value?.trim() || "";
  }

  async function fetchHtml(url) {
    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      cache: "no-cache",
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${url}`);
    }
    return response.text();
  }

  async function ensureToken() {
    if (cachedToken) return cachedToken;

    const pageToken = document
      .querySelector('input[name="token"]')
      ?.value?.trim();
    if (pageToken) {
      cachedToken = pageToken;
      return cachedToken;
    }

    const html = await fetchHtml(SEARCH_URL);
    const token = parseTokenFromHtml(html);
    if (!token) throw new Error("未获取到 token");
    cachedToken = token;
    return token;
  }

  function extractArrayLiteralFrom(scriptText, fromPos) {
    const start = scriptText.indexOf("[", fromPos);
    if (start < 0) return "";

    let i = start;
    let depth = 0;
    let quote = "";
    let escaped = false;
    for (; i < scriptText.length; i += 1) {
      const ch = scriptText[i];
      if (quote) {
        if (escaped) escaped = false;
        else if (ch === "\\") escaped = true;
        else if (ch === quote) quote = "";
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        continue;
      }
      if (ch === "[") {
        depth += 1;
        continue;
      }
      if (ch === "]") {
        depth -= 1;
        if (depth === 0) return scriptText.slice(start, i + 1);
      }
    }
    return "";
  }

  function escapeRegExp(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function extractGridData(html, gridId) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const scripts = Array.from(doc.querySelectorAll("script"));

    const selectorRe = new RegExp(
      `\\$\\(\\s*["']#${escapeRegExp(gridId)}["']\\s*\\)\\s*\\.kendoGrid\\s*\\(`,
    );
    for (const script of scripts) {
      const text = script.textContent || "";
      const selectorMatch = selectorRe.exec(text);
      if (!selectorMatch) continue;

      const fromSelector = selectorMatch.index;
      const dataKeyPos = text.indexOf("data:", fromSelector);
      if (dataKeyPos < 0) continue;

      const literal = extractArrayLiteralFrom(text, dataKeyPos);
      if (!literal) continue;

      let parsed = safeJsonParse(literal, null);
      if (!parsed) parsed = safeEvalLiteral(literal, null);
      if (Array.isArray(parsed)) return parsed;
    }
    return [];
  }

  function extractHerbMeta(html) {
    const match = html.match(/var\s+data_hname\s*=\s*(\{[\s\S]*?\});/);
    if (!match) return {};
    return safeJsonParse(match[1], safeEvalLiteral(match[1], {})) || {};
  }

  function chooseJoinKey(ingredients, targets, diseases) {
    const candidates = ["MOL_ID", "molecule_ID", "molecule_name"];
    const iKeys = new Set(Object.keys(ingredients[0] || {}));
    const tKeys = new Set(Object.keys(targets[0] || {}));
    const dKeys = new Set(Object.keys(diseases[0] || {}));
    for (const key of candidates) {
      if (iKeys.has(key) && (tKeys.has(key) || dKeys.has(key))) return key;
    }
    return "";
  }

  function mergeDetailRows(ingredients, targets, diseases) {
    const joinKey = chooseJoinKey(ingredients, targets, diseases);
    if (!joinKey) {
      return ingredients.map((item) => ({
        ...item,
        related_targets: "",
        related_diseases: "",
        target_count: 0,
        disease_count: 0,
      }));
    }

    const targetMap = new Map();
    for (const row of targets) {
      const key = row[joinKey];
      if (!key) continue;
      if (!targetMap.has(key)) targetMap.set(key, []);
      targetMap.get(key).push(row);
    }

    const diseaseMap = new Map();
    for (const row of diseases) {
      const key = row[joinKey];
      if (!key) continue;
      if (!diseaseMap.has(key)) diseaseMap.set(key, []);
      diseaseMap.get(key).push(row);
    }

    return ingredients.map((item) => {
      const key = item[joinKey];
      const matchedTargets = key ? targetMap.get(key) || [] : [];
      const matchedDiseases = key ? diseaseMap.get(key) || [] : [];
      return {
        ...item,
        related_targets: matchedTargets
          .map((x) => x.target_name || x.TARGET_ID || "")
          .filter(Boolean)
          .join("; "),
        related_diseases: matchedDiseases
          .map((x) => x.disease_name || x.disease_ID || "")
          .filter(Boolean)
          .join("; "),
        target_count: matchedTargets.length,
        disease_count: matchedDiseases.length,
      };
    });
  }

  function sanitizeName(name) {
    return String(name || "")
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, "_")
      .slice(0, 80);
  }

  function setStatus(text, isError) {
    const node = document.querySelector(`#${panelId} .tcmsp-status`);
    if (!node) return;
    node.textContent = text;
    node.style.color = isError ? "#b91c1c" : "#111827";
  }

  function withMarkers(rows, context) {
    const {
      queryKeyword,
      queryIndex,
      herbIndex,
      herbCn,
      herbPinyin,
      herbEn,
      herbCode,
    } = context;
    return rows.map((row) => ({
      query_keyword: queryKeyword,
      query_index: queryIndex,
      herb_index: herbIndex,
      herb_code: herbCode,
      herb_cn_name: herbCn,
      herb_pinyin: herbPinyin,
      herb_en_name: herbEn,
      ...row,
    }));
  }

  async function processSingleHerb(
    token,
    herb,
    queryKeyword,
    queryIndex,
    herbIndex,
    totalHerbs,
    report,
    filterOptions,
  ) {
    const herbEn = herb.herb_en_name || herb.herb_cn_name || "";
    report(`抓取详情 [${queryIndex}] ${herbIndex}/${totalHerbs}: ${herbEn}`);

    const detailHtml = await fetchHtml(
      `${SEARCH_URL}?qr=${encodeURIComponent(herbEn)}&qsr=herb_en_name&token=${encodeURIComponent(token)}`,
    );

    const herbMeta = { ...herb, ...extractHerbMeta(detailHtml) };
    const ingredients = extractGridData(detailHtml, "grid");
    const targets = extractGridData(detailHtml, "grid2");
    const diseases = extractGridData(detailHtml, "grid3");
    const filteredData = applyIngredientFilters(
      ingredients,
      targets,
      diseases,
      filterOptions,
    );
    const merged = mergeDetailRows(
      filteredData.ingredients,
      filteredData.targets,
      filteredData.diseases,
    );

    const herbCode = sanitizeName(
      `${herbMeta.herb_pinyin || herbMeta.herb_en_name || herbMeta.herb_cn_name || herbIndex}`,
    );
    const marker = {
      queryKeyword,
      queryIndex,
      herbIndex,
      herbCn: herbMeta.herb_cn_name || "",
      herbPinyin: herbMeta.herb_pinyin || "",
      herbEn: herbMeta.herb_en_name || herbEn,
      herbCode,
    };

    return {
      herb: herbMeta,
      counts: {
        ingredients: filteredData.ingredients.length,
        targets: filteredData.targets.length,
        diseases: filteredData.diseases.length,
      },
      ingredientsRows: withMarkers(filteredData.ingredients, marker),
      targetsRows: withMarkers(filteredData.targets, marker),
      diseasesRows: withMarkers(filteredData.diseases, marker),
      mergedRows: withMarkers(merged, marker),
    };
  }

  async function processQuery(
    token,
    queryKeyword,
    queryIndex,
    report,
    filterOptions,
  ) {
    report(`检索药物 [${queryIndex}]: ${queryKeyword}`);

    const searchHtml = await fetchHtml(
      `${SEARCH_URL}?qs=herb_all_name&q=${encodeURIComponent(queryKeyword)}&token=${encodeURIComponent(token)}`,
    );
    const herbs = extractGridData(searchHtml, "grid");
    if (!herbs.length) {
      return {
        queryKeyword,
        queryIndex,
        found: false,
        herbs: [],
        ingredientsRows: [],
        targetsRows: [],
        diseasesRows: [],
        mergedRows: [],
      };
    }

    const herbResults = [];
    for (let i = 0; i < herbs.length; i += 1) {
      const item = await processSingleHerb(
        token,
        herbs[i],
        queryKeyword,
        queryIndex,
        i + 1,
        herbs.length,
        report,
        filterOptions,
      );
      herbResults.push(item);
    }

    return {
      queryKeyword,
      queryIndex,
      found: true,
      herbs: herbResults.map((h) => h.herb),
      ingredientsRows: herbResults.flatMap((h) => h.ingredientsRows),
      targetsRows: herbResults.flatMap((h) => h.targetsRows),
      diseasesRows: herbResults.flatMap((h) => h.diseasesRows),
      mergedRows: herbResults.flatMap((h) => h.mergedRows),
      stats: herbResults.map((h) => ({
        herb_cn_name: h.herb.herb_cn_name || "",
        herb_pinyin: h.herb.herb_pinyin || "",
        herb_en_name: h.herb.herb_en_name || "",
        ingredients: h.counts.ingredients,
        targets: h.counts.targets,
        diseases: h.counts.diseases,
      })),
    };
  }

  async function runWithConcurrency(items, limit, worker) {
    const results = new Array(items.length);
    let cursor = 0;

    async function runner() {
      while (true) {
        const idx = cursor;
        cursor += 1;
        if (idx >= items.length) return;
        results[idx] = await worker(items[idx], idx);
      }
    }

    const size = Math.min(limit, items.length);
    const tasks = [];
    for (let i = 0; i < size; i += 1) tasks.push(runner());
    await Promise.all(tasks);
    return results;
  }

  function sheetFromRows(rows) {
    if (!rows.length) return XLSX.utils.aoa_to_sheet([["no_data"]]);
    return XLSX.utils.json_to_sheet(rows);
  }

  function safeSheetName(name, fallback) {
    const raw = sanitizeName(name || "");
    const valid = raw.replace(/[\[\]\*\/\\\?:]/g, "_").slice(0, 31);
    return valid || fallback;
  }

  function buildWorkbook(queryResults, mergeAll) {
    const wb = XLSX.utils.book_new();
    const summaryRows = queryResults.map((qr) => ({
      query_index: qr.queryIndex,
      query_keyword: qr.queryKeyword,
      status: qr.found ? "FOUND" : "NOT_FOUND",
      matched_herb_count: qr.herbs.length,
      ingredients_rows: qr.ingredientsRows.length,
      targets_rows: qr.targetsRows.length,
      diseases_rows: qr.diseasesRows.length,
      merged_rows: qr.mergedRows.length,
    }));
    XLSX.utils.book_append_sheet(wb, sheetFromRows(summaryRows), "summary");

    if (mergeAll) {
      const ingredients = queryResults.flatMap((x) => x.ingredientsRows);
      const targets = queryResults.flatMap((x) => x.targetsRows);
      const diseases = queryResults.flatMap((x) => x.diseasesRows);
      const merged = queryResults.flatMap((x) => x.mergedRows);

      XLSX.utils.book_append_sheet(wb, sheetFromRows(merged), "merged_all");
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows(ingredients),
        "ingredients_all",
      );
      XLSX.utils.book_append_sheet(wb, sheetFromRows(targets), "targets_all");
      XLSX.utils.book_append_sheet(wb, sheetFromRows(diseases), "diseases_all");
      return wb;
    }

    for (const qr of queryResults) {
      const prefix = `q${qr.queryIndex}_${qr.queryKeyword}`;
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows(qr.mergedRows),
        safeSheetName(`${prefix}_merged`, `q${qr.queryIndex}_merged`),
      );
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows(qr.ingredientsRows),
        safeSheetName(`${prefix}_ingredients`, `q${qr.queryIndex}_ingredients`),
      );
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows(qr.targetsRows),
        safeSheetName(`${prefix}_targets`, `q${qr.queryIndex}_targets`),
      );
      XLSX.utils.book_append_sheet(
        wb,
        sheetFromRows(qr.diseasesRows),
        safeSheetName(`${prefix}_diseases`, `q${qr.queryIndex}_diseases`),
      );
    }
    return wb;
  }

  function downloadWorkbook(workbook, filename) {
    XLSX.writeFile(workbook, filename, { compression: true });
  }

  function parseQueries(text) {
    return Array.from(
      new Set(
        String(text || "")
          .split(/\r?\n/)
          .map((x) => x.trim())
          .filter(Boolean),
      ),
    );
  }

  function parseNumberish(value) {
    const num = Number.parseFloat(String(value ?? "").trim());
    return Number.isFinite(num) ? num : Number.NaN;
  }

  function applyIngredientFilters(
    ingredients,
    targets,
    diseases,
    filterOptions,
  ) {
    const obEnabled = Boolean(filterOptions?.obEnabled);
    const dlEnabled = Boolean(filterOptions?.dlEnabled);
    const obMin = Number.isFinite(filterOptions?.obMin)
      ? filterOptions.obMin
      : 30;
    const dlMin = Number.isFinite(filterOptions?.dlMin)
      ? filterOptions.dlMin
      : 0.18;

    if (!obEnabled && !dlEnabled) {
      return { ingredients, targets, diseases };
    }

    const filteredIngredients = ingredients.filter((row) => {
      if (obEnabled) {
        const ob = parseNumberish(row?.ob);
        if (!Number.isFinite(ob) || ob < obMin) return false;
      }
      if (dlEnabled) {
        const dl = parseNumberish(row?.dl);
        if (!Number.isFinite(dl) || dl < dlMin) return false;
      }
      return true;
    });

    const joinKey = chooseJoinKey(ingredients, targets, diseases);
    if (!joinKey) {
      return {
        ingredients: filteredIngredients,
        targets,
        diseases,
      };
    }

    const keptKeys = new Set(
      filteredIngredients.map((row) => row?.[joinKey]).filter(Boolean),
    );
    return {
      ingredients: filteredIngredients,
      targets: targets.filter((row) => keptKeys.has(row?.[joinKey])),
      diseases: diseases.filter((row) => keptKeys.has(row?.[joinKey])),
    };
  }

  async function queryAndDownloadBatch(queryKeywords, mergeAll, filterOptions) {
    if (typeof XLSX === "undefined") {
      throw new Error("XLSX 库未加载，请刷新后重试");
    }

    const token = await ensureToken();
    setStatus(`开始处理 ${queryKeywords.length} 个查询词...`);

    const queryResults = await runWithConcurrency(
      queryKeywords,
      QUERY_CONCURRENCY,
      async (query, idx) =>
        processQuery(token, query, idx + 1, setStatus, filterOptions),
    );

    const now = new Date().toISOString();
    const fileKeyword =
      queryKeywords.length === 1
        ? queryKeywords[0]
        : `batch_${queryKeywords.length}_items`;
    const fileBase = `TCMSP_${sanitizeName(fileKeyword)}_${now.slice(0, 19).replace(/[:T]/g, "-")}`;
    const workbook = buildWorkbook(queryResults, mergeAll);
    downloadWorkbook(workbook, `${fileBase}.xlsx`);

    const notFound = queryResults
      .filter((x) => !x.found)
      .map((x) => x.queryKeyword);
    const okCount = queryResults.length - notFound.length;
    if (!notFound.length) {
      setStatus(
        `完成：${okCount}/${queryResults.length} 查询成功，已下载 XLSX`,
      );
    } else {
      setStatus(
        `完成：成功 ${okCount}，未命中 ${notFound.length}（${notFound.join("、")}）`,
        true,
      );
    }
  }

  function createPanel() {
    if (document.getElementById(panelId)) return;

    const style = document.createElement("style");
    style.textContent = `
      #${panelId}{
        position:fixed;right:16px;top:72px;z-index:999999;
        width:370px;padding:10px 14px 14px;border-radius:12px;
        border:1px solid #d1d5db;background:#ffffff;box-shadow:0 10px 30px rgba(0,0,0,.14);
        font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
        box-sizing:border-box;
      }
      #${panelId} .tcmsp-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px;cursor:move;}
      #${panelId} .tcmsp-title{font-size:16px;font-weight:700;color:#111827;margin:0;}
      #${panelId} .tcmsp-collapse{
        width:26px;height:26px;border-radius:6px;border:1px solid #cbd5e1;background:#f8fafc;
        color:#334155;font-size:16px;line-height:1;cursor:pointer;
      }
      #${panelId} .tcmsp-body{display:block;}
      #${panelId}.is-collapsed .tcmsp-body{display:none;}
      #${panelId} .tcmsp-input{
        width:100%;min-height:110px;padding:10px 12px;border-radius:10px;
        border:1px solid #cbd5e1;color:#111827;font-size:13px;line-height:1.45;outline:none;resize:vertical;
        box-sizing:border-box;
      }
      #${panelId} .tcmsp-input:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.16);}
      #${panelId} .tcmsp-options{margin-top:8px;display:flex;align-items:center;justify-content:space-between;gap:8px;}
      #${panelId} .tcmsp-check{font-size:12px;color:#374151;display:flex;align-items:center;gap:6px;user-select:none;}
      #${panelId} .tcmsp-filters{margin-top:8px;border:1px solid #e5e7eb;border-radius:10px;padding:8px;background:#f9fafb;}
      #${panelId} .tcmsp-filter-row{display:flex;align-items:center;justify-content:space-between;gap:8px;}
      #${panelId} .tcmsp-filter-row + .tcmsp-filter-row{margin-top:6px;}
      #${panelId} .tcmsp-filter-num{
        width:92px;height:28px;padding:0 8px;border-radius:8px;border:1px solid #cbd5e1;
        font-size:12px;color:#111827;box-sizing:border-box;
      }
      #${panelId} .tcmsp-btn{
        height:34px;padding:0 14px;border-radius:8px;border:1px solid #2563eb;
        background:#2563eb;color:#fff;font-size:12px;cursor:pointer;white-space:nowrap;
      }
      #${panelId} .tcmsp-btn[disabled]{opacity:.6;cursor:not-allowed;}
      #${panelId} .tcmsp-status{margin-top:10px;font-size:12px;line-height:1.45;color:#111827;white-space:pre-wrap;}
      #${panelId} .tcmsp-tip{margin-top:6px;font-size:11px;line-height:1.4;color:#6b7280;}
      @media (max-width: 680px){
        #${panelId}{left:10px;right:10px;top:10px;width:auto;}
        #${panelId} .tcmsp-head{cursor:default;}
      }
    `;
    document.head.appendChild(style);

    const panel = document.createElement("div");
    panel.id = panelId;
    panel.innerHTML = `
      <div class="tcmsp-head">
        <p class="tcmsp-title">TCMSP Exporter</p>
        <button class="tcmsp-collapse" type="button" title="折叠/展开">−</button>
      </div>
      <div class="tcmsp-body">
        <textarea class="tcmsp-input" placeholder="每行一个药物名（支持中文/拼音/拉丁名），例如：
陈皮
麻黄
Baizhu
Citrus Reticulata"></textarea>
        <div class="tcmsp-filters">
          <div class="tcmsp-filter-row">
            <label class="tcmsp-check"><input type="checkbox" class="tcmsp-ob-enabled" checked> OB (%) ≥</label>
            <input class="tcmsp-filter-num tcmsp-ob-min" type="number" min="0" step="0.1" value="30">
          </div>
          <div class="tcmsp-filter-row">
            <label class="tcmsp-check"><input type="checkbox" class="tcmsp-dl-enabled" checked> DL ≥</label>
            <input class="tcmsp-filter-num tcmsp-dl-min" type="number" min="0" step="0.01" value="0.18">
          </div>
        </div>
        <div class="tcmsp-options">
          <label class="tcmsp-check"><input type="checkbox" class="tcmsp-merge" checked> 合并多药物结果到同一表</label>
          <button class="tcmsp-btn" type="button">抓取并下载 XLSX</button>
        </div>
        <div class="tcmsp-status">就绪</div>
        <div class="tcmsp-tip">将下载一个 XLSX 文件：summary + merged / ingredients / targets / diseases（按合并选项输出；可按 OB/DL 阈值筛选）</div>
      </div>
    `;
    document.body.appendChild(panel);

    const head = panel.querySelector(".tcmsp-head");
    const collapseBtn = panel.querySelector(".tcmsp-collapse");
    const input = panel.querySelector(".tcmsp-input");
    const mergeCheckbox = panel.querySelector(".tcmsp-merge");
    const obEnabledCheckbox = panel.querySelector(".tcmsp-ob-enabled");
    const dlEnabledCheckbox = panel.querySelector(".tcmsp-dl-enabled");
    const obMinInput = panel.querySelector(".tcmsp-ob-min");
    const dlMinInput = panel.querySelector(".tcmsp-dl-min");
    const button = panel.querySelector(".tcmsp-btn");

    function savePanelPos(left, top) {
      localStorage.setItem(POS_STORE_KEY, JSON.stringify({ left, top }));
    }

    function loadPanelPos() {
      const raw = localStorage.getItem(POS_STORE_KEY);
      const pos = safeJsonParse(raw, null);
      if (!pos || typeof pos.left !== "number" || typeof pos.top !== "number")
        return;
      panel.style.left = `${pos.left}px`;
      panel.style.top = `${pos.top}px`;
      panel.style.right = "auto";
    }

    function applyCollapsed(collapsed) {
      panel.classList.toggle("is-collapsed", collapsed);
      collapseBtn.textContent = collapsed ? "+" : "−";
      collapseBtn.title = collapsed ? "展开" : "折叠";
      localStorage.setItem(COLLAPSE_STORE_KEY, collapsed ? "1" : "0");
    }

    function loadCollapsed() {
      applyCollapsed(localStorage.getItem(COLLAPSE_STORE_KEY) === "1");
    }

    function makeDraggable() {
      let dragging = false;
      let offsetX = 0;
      let offsetY = 0;

      head.addEventListener("mousedown", (event) => {
        if (event.target === collapseBtn) return;
        if (window.matchMedia("(max-width: 680px)").matches) return;
        dragging = true;
        const rect = panel.getBoundingClientRect();
        offsetX = event.clientX - rect.left;
        offsetY = event.clientY - rect.top;
        panel.style.left = `${rect.left}px`;
        panel.style.top = `${rect.top}px`;
        panel.style.right = "auto";
        document.body.style.userSelect = "none";
        event.preventDefault();
      });

      window.addEventListener("mousemove", (event) => {
        if (!dragging) return;
        const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
        const maxTop = Math.max(0, window.innerHeight - 40);
        const left = Math.min(Math.max(0, event.clientX - offsetX), maxLeft);
        const top = Math.min(Math.max(0, event.clientY - offsetY), maxTop);
        panel.style.left = `${left}px`;
        panel.style.top = `${top}px`;
      });

      window.addEventListener("mouseup", () => {
        if (!dragging) return;
        dragging = false;
        document.body.style.userSelect = "";
        const left = parseFloat(panel.style.left || "0");
        const top = parseFloat(panel.style.top || "0");
        savePanelPos(left, top);
      });
    }

    async function run() {
      const queries = parseQueries(input.value);
      if (!queries.length) {
        setStatus("请输入药物名，每行一个", true);
        return;
      }

      button.disabled = true;
      try {
        const obMin = parseNumberish(obMinInput.value);
        const dlMin = parseNumberish(dlMinInput.value);
        const filterOptions = {
          obEnabled: Boolean(obEnabledCheckbox.checked),
          dlEnabled: Boolean(dlEnabledCheckbox.checked),
          obMin: Number.isFinite(obMin) ? obMin : 30,
          dlMin: Number.isFinite(dlMin) ? dlMin : 0.18,
        };
        await queryAndDownloadBatch(
          queries,
          Boolean(mergeCheckbox.checked),
          filterOptions,
        );
      } catch (err) {
        setStatus(`失败：${err?.message || err}`, true);
      } finally {
        button.disabled = false;
      }
    }

    button.addEventListener("click", run);
    input.addEventListener("keydown", (event) => {
      if ((event.metaKey || event.ctrlKey) && event.key === "Enter") run();
    });
    collapseBtn.addEventListener("click", () => {
      applyCollapsed(!panel.classList.contains("is-collapsed"));
    });

    loadPanelPos();
    loadCollapsed();
    makeDraggable();
  }

  function boot() {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", createPanel, {
        once: true,
      });
    } else {
      createPanel();
    }
  }

  boot();
})();
