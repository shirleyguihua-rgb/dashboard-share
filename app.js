const REQUIRED_SHEETS = {
  profit: "领星 ERP ASIN 维度利润统计报表",
  plan: "Q2销售预期",
};

const IS_STANDALONE =
  window.__STANDALONE__ === true || window.location.protocol === "file:";
const IS_READONLY = window.__READONLY__ === true;

const state = {
  sourceMode: "excel",
  dataset: null,
  manualDate: "",
  detailFilters: { market: "", country: "", asin: "", product: "", owner: "" },
  weeklyFilters: { market: "", country: "", asin: "", product: "", owner: "" },
  monthlyFilters: { month: "", market: "", country: "", asin: "", owner: "" },
  warningFilters: { risk: "", market: "", country: "", asin: "", product: "", owner: "" },
  countrySort: { key: "latest", direction: "desc" },
  warningSort: { key: "riskScore", direction: "desc" },
  detailSort: { key: "latest", direction: "desc" },
  weeklySort: { key: "completion", direction: "desc" },
  monthlySort: { key: "completion", direction: "desc" },
};

const refs = {
  sourceStatus: document.getElementById("sourceStatus"),
  excelPanel: document.getElementById("excelPanel"),
  feishuPanel: document.getElementById("feishuPanel"),
  excelFile: document.getElementById("excelFile"),
  syncFeishuBtn: document.getElementById("syncFeishuBtn"),
  baseUrl: document.getElementById("baseUrl"),
  appId: document.getElementById("appId"),
  appSecret: document.getElementById("appSecret"),
  appToken: document.getElementById("appToken"),
  profitTableName: document.getElementById("profitTableName"),
  planTableName: document.getElementById("planTableName"),
  manualDate: document.getElementById("manualDate"),
  applyDateBtn: document.getElementById("applyDateBtn"),
  resetDateBtn: document.getElementById("resetDateBtn"),
  publishShareBtn: document.getElementById("publishShareBtn"),
  shareStatus: document.getElementById("shareStatus"),
  shareLink: document.getElementById("shareLink"),
  heroMeta: document.getElementById("heroMeta"),
  heroPill: document.getElementById("heroPill"),
  kpiGrid: document.getElementById("kpiGrid"),
  summaryCount: document.getElementById("summaryCount"),
  warningCount: document.getElementById("warningCount"),
  warningTableCount: document.getElementById("warningTableCount"),
  warningRiskFilter: document.getElementById("warningRiskFilter"),
  warningMarketFilter: document.getElementById("warningMarketFilter"),
  warningCountryFilter: document.getElementById("warningCountryFilter"),
  warningAsinFilter: document.getElementById("warningAsinFilter"),
  warningProductFilter: document.getElementById("warningProductFilter"),
  warningOwnerFilter: document.getElementById("warningOwnerFilter"),
  detailCount: document.getElementById("detailCount"),
  weeklyCount: document.getElementById("weeklyCount"),
  monthlyCount: document.getElementById("monthlyCount"),
  countryBars: document.getElementById("countryBars"),
  countryTableBody: document.getElementById("countryTableBody"),
  warningTopCards: document.getElementById("warningTopCards"),
  warningTableBody: document.getElementById("warningTableBody"),
  detailMarketFilter: document.getElementById("detailMarketFilter"),
  detailCountryFilter: document.getElementById("detailCountryFilter"),
  detailAsinFilter: document.getElementById("detailAsinFilter"),
  detailProductFilter: document.getElementById("detailProductFilter"),
  detailOwnerFilter: document.getElementById("detailOwnerFilter"),
  detailTableBody: document.getElementById("detailTableBody"),
  weeklyMarketFilter: document.getElementById("weeklyMarketFilter"),
  weeklyCountryFilter: document.getElementById("weeklyCountryFilter"),
  weeklyAsinFilter: document.getElementById("weeklyAsinFilter"),
  weeklyProductFilter: document.getElementById("weeklyProductFilter"),
  weeklyOwnerFilter: document.getElementById("weeklyOwnerFilter"),
  weeklyTableBody: document.getElementById("weeklyTableBody"),
  monthlyMonthFilter: document.getElementById("monthlyMonthFilter"),
  monthlyMarketFilter: document.getElementById("monthlyMarketFilter"),
  monthlyCountryFilter: document.getElementById("monthlyCountryFilter"),
  monthlyAsinFilter: document.getElementById("monthlyAsinFilter"),
  monthlyOwnerFilter: document.getElementById("monthlyOwnerFilter"),
  monthlyTableBody: document.getElementById("monthlyTableBody"),
  toast: document.getElementById("toast"),
};

[
  ["detailCountryFilter", "detailFilters", "country"],
  ["detailMarketFilter", "detailFilters", "market"],
  ["detailAsinFilter", "detailFilters", "asin"],
  ["detailProductFilter", "detailFilters", "product"],
  ["detailOwnerFilter", "detailFilters", "owner"],
  ["warningRiskFilter", "warningFilters", "risk"],
  ["warningMarketFilter", "warningFilters", "market"],
  ["warningCountryFilter", "warningFilters", "country"],
  ["warningAsinFilter", "warningFilters", "asin"],
  ["warningProductFilter", "warningFilters", "product"],
  ["warningOwnerFilter", "warningFilters", "owner"],
  ["weeklyMarketFilter", "weeklyFilters", "market"],
  ["weeklyCountryFilter", "weeklyFilters", "country"],
  ["weeklyAsinFilter", "weeklyFilters", "asin"],
  ["weeklyProductFilter", "weeklyFilters", "product"],
  ["weeklyOwnerFilter", "weeklyFilters", "owner"],
  ["monthlyMonthFilter", "monthlyFilters", "month"],
  ["monthlyMarketFilter", "monthlyFilters", "market"],
  ["monthlyCountryFilter", "monthlyFilters", "country"],
  ["monthlyAsinFilter", "monthlyFilters", "asin"],
  ["monthlyOwnerFilter", "monthlyFilters", "owner"],
].forEach(([refKey, stateKey, field]) => {
  refs[refKey].addEventListener("change", (event) => {
    state[stateKey][field] = event.target.value;
    renderDashboard();
  });
});

document.querySelectorAll("th.sortable").forEach((th) => {
  th.addEventListener("click", () => {
    const table = th.dataset.sortTable;
    const key = th.dataset.sortKey;
    const stateKey =
      table === "country"
        ? "countrySort"
        : table === "warning"
          ? "warningSort"
        : table === "detail"
          ? "detailSort"
          : table === "weekly"
            ? "weeklySort"
            : "monthlySort";
    const current = state[stateKey];
    state[stateKey] = {
      key,
      direction:
        current.key === key && current.direction === "desc" ? "asc" : "desc",
    };
    renderDashboard();
  });
});

document.querySelectorAll('input[name="sourceMode"]').forEach((input) => {
  input.addEventListener("change", () => {
    if (IS_STANDALONE && input.value === "feishu") {
      input.checked = false;
      document.querySelector('input[name="sourceMode"][value="excel"]').checked = true;
      state.sourceMode = "excel";
      refs.excelPanel.classList.remove("hidden");
      refs.feishuPanel.classList.add("hidden");
      showToast("离线单文件版仅支持上传 Excel，飞书直连需要本地服务模式");
      return;
    }
    state.sourceMode = input.value;
    refs.excelPanel.classList.toggle("hidden", input.value !== "excel");
    refs.feishuPanel.classList.toggle("hidden", input.value !== "feishu");
  });
});

if (refs.excelFile) refs.excelFile.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    setStatus("处理中", "badge-info");
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: true,
      raw: false,
    });
    const dataset = buildDatasetFromWorkbook(workbook);
    state.dataset = dataset;
    hydrateDateInput();
    renderDashboard();
    setStatus(`已加载 ${file.name}`, "badge-success");
    showToast(`已从 ${file.name} 加载数据`);
  } catch (error) {
    console.error(error);
    setStatus("加载失败", "badge-danger");
    showToast(`Excel 解析失败：${error.message}`);
  }
});

if (refs.syncFeishuBtn) refs.syncFeishuBtn.addEventListener("click", async () => {
  if (IS_STANDALONE) {
    showToast("离线单文件版不支持飞书直连，请使用 Excel 上传更新数据");
    return;
  }

  const payload = {
    baseUrl: refs.baseUrl.value.trim(),
    appId: refs.appId.value.trim(),
    appSecret: refs.appSecret.value.trim(),
    appToken: refs.appToken.value.trim(),
    profitTableName: refs.profitTableName.value.trim(),
    planTableName: refs.planTableName.value.trim(),
  };

  if (!payload.baseUrl) {
    showToast("请先填写飞书 Base 链接");
    return;
  }

  try {
    setStatus("同步中", "badge-info");
    const endpoint =
      payload.appId && payload.appSecret && payload.appToken
        ? "/api/feishu/dashboard"
        : "/api/feishu/session-dashboard";

    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error || "飞书同步失败");
    }

    state.dataset = buildDatasetFromNormalizedRows(result);
    hydrateDateInput();
    renderDashboard();
    setStatus(
      endpoint.includes("session") ? "飞书已同步（登录态）" : "飞书已同步（OpenAPI）",
      "badge-success"
    );
    showToast(
      endpoint.includes("session")
        ? "已通过浏览器登录态同步飞书数据"
        : "已通过 OpenAPI 同步飞书数据"
    );
  } catch (error) {
    console.error(error);
    setStatus("同步失败", "badge-danger");
    showToast(error.message);
  }
});

if (refs.applyDateBtn) refs.applyDateBtn.addEventListener("click", () => {
  state.manualDate = refs.manualDate.value;
  renderDashboard();
});

if (refs.resetDateBtn) refs.resetDateBtn.addEventListener("click", () => {
  state.manualDate = "";
  if (refs.manualDate) refs.manualDate.value = "";
  renderDashboard();
});

if (refs.publishShareBtn) {
  refs.publishShareBtn.addEventListener("click", async () => {
    if (!state.dataset) {
      showToast("请先同步飞书或上传 Excel，再发布分享版");
      return;
    }
    try {
      refs.shareStatus.textContent = "发布中";
      refs.shareStatus.className = "badge badge-info";
      const response = await fetch("/api/share/publish", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dataset: state.dataset,
          publishedAt: new Date().toISOString(),
        }),
      });
      const result = await response.json();
      if (!response.ok) {
        throw new Error(result.error || "分享版发布失败");
      }
      refs.shareStatus.textContent = "已发布";
      refs.shareStatus.className = "badge badge-success";
      refs.shareLink.value = result.localUrl || "";
      showToast("分享版已生成，可复制只读链接或上传 share_bundle 目录");
    } catch (error) {
      console.error(error);
      refs.shareStatus.textContent = "发布失败";
      refs.shareStatus.className = "badge badge-danger";
      showToast(error.message);
    }
  });
}

function buildDatasetFromWorkbook(workbook) {
  const profitSheet = workbook.Sheets[REQUIRED_SHEETS.profit];
  const planSheet = workbook.Sheets[REQUIRED_SHEETS.plan];
  const ownerSheet =
    workbook.Sheets["SKU映射对应表"] || workbook.Sheets["日毛利汇总看板"];
  if (!profitSheet || !planSheet) {
    throw new Error(
      `工作簿中缺少必要工作表，请确认存在 ${REQUIRED_SHEETS.profit} 和 ${REQUIRED_SHEETS.plan}`
    );
  }

  const profitRows = normalizeProfitRows(
    XLSX.utils.sheet_to_json(profitSheet, {
      defval: null,
      raw: false,
    })
  );
  const planRows = normalizePlanRows(
    XLSX.utils.sheet_to_json(planSheet, {
      defval: null,
      raw: false,
    })
  );
  const ownerRows = ownerSheet
    ? normalizeOwnerRows(
        XLSX.utils.sheet_to_json(ownerSheet, {
          defval: null,
          raw: false,
        })
      )
    : [];
  const weeklyTrackSheet = workbook.Sheets["Q2周进度追踪"];
  const monthlyTrackSheet = workbook.Sheets["Q2月进度追踪"];
  const weeklyTrackRows = weeklyTrackSheet
    ? normalizeWeeklyTrackRows(
        XLSX.utils.sheet_to_json(weeklyTrackSheet, {
          defval: null,
          raw: false,
        })
      )
    : [];
  const monthlyTrackRows = monthlyTrackSheet
    ? normalizeMonthlyTrackRows(
        XLSX.utils.sheet_to_json(monthlyTrackSheet, {
          defval: null,
          raw: false,
        })
      )
    : [];

  return {
    profitRows,
    planRows,
    ownerRows,
    weeklyTrackRows,
    monthlyTrackRows,
    sourceMeta: { source: "excel" },
  };
}

function buildDatasetFromNormalizedRows(payload) {
  return {
    profitRows: normalizeProfitRows(payload.profitRows || []),
    planRows: normalizePlanRows(payload.planRows || []),
    ownerRows: normalizeOwnerRows(payload.ownerRows || []),
    weeklyTrackRows: normalizeWeeklyTrackRows(payload.weeklyTrackRows || []),
    monthlyTrackRows: normalizeMonthlyTrackRows(payload.monthlyTrackRows || []),
    sourceMeta: {
      source: "feishu",
      updatedAt: payload.updatedAt || new Date().toISOString(),
    },
  };
}

function normalizeProfitRows(rows) {
  return rows
    .map((row) => ({
      date: normalizeDate(
        row["日期"] ?? row.date ?? row["date"] ?? row["Date"] ?? null
      ),
      country: normalizeText(row["国家"] ?? row.country),
      asin: normalizeText(row["asin"] ?? row["ASIN"] ?? row.asin),
      product: normalizeText(row["品名"] ?? row.product),
      grossProfitRmb: normalizeNumber(
        row["毛利润（RMB）"] ??
          row["毛利润RMB"] ??
          row["grossProfitRmb"] ??
          row["gross_profit_rmb"]
      ),
    }))
    .filter((row) => row.date && row.country);
}

function normalizePlanRows(rows) {
  return rows
    .map((row) => {
      const weekRange = normalizeText(row["周区间"] ?? row.weekRange);
      const parsedRange = parseWeekRange(weekRange);
      return {
        weekNo: normalizeInteger(row["周序号"] ?? row.weekNo),
        weekRange,
        weekStart: parsedRange.start,
        weekEnd:
          normalizeDate(row["周结束日"] ?? row.weekEnd) ?? parsedRange.end,
        country: normalizeText(row["国家"] ?? row.country),
        asin: normalizeText(row["ASIN"] ?? row["asin"] ?? row.asin),
        product: normalizeText(row["品名"] ?? row.product),
        marketLevel: normalizeText(row["市场等级"] ?? row.marketLevel),
        listingLevel: normalizeText(row["链接分级"] ?? row.listingLevel),
        weeklyProfitRmb: normalizeNumber(
          row["周毛利（RMB）"] ?? row["weeklyProfitRmb"]
        ),
        dailyProfitRmb: normalizeNumber(
          row["单日毛利（RMB）"] ?? row["dailyProfitRmb"]
        ),
      };
    })
    .filter((row) => row.country && row.asin);
}

function normalizeOwnerRows(rows) {
  return rows
    .map((row) => ({
      country: normalizeText(row["国家"] ?? row.country),
      asin: normalizeText(row["ASIN"] ?? row["asin"] ?? row.asin),
      product: normalizeText(row["品名"] ?? row.product),
      owner: normalizeText(row["负责人"] ?? row.owner ?? row["文本 5"]),
    }))
    .filter((row) => row.country && row.asin);
}

function normalizeWeeklyTrackRows(rows) {
  return rows
    .map((row) => ({
      country: normalizeText(row["国家"] ?? row.country),
      asin: normalizeText(row["ASIN"] ?? row["asin"] ?? row.asin),
      product: normalizeText(row["品名"] ?? row.product),
      marketLevel: normalizeText(row["市场等级"] ?? row.marketLevel),
      listingLevel: normalizeText(row["链接分级"] ?? row.listingLevel),
      owner: normalizeText(row["负责人"] ?? row.owner ?? row["文本 5"]),
      weeklyPlan: normalizeNumber(row["周毛利预期RMB"] ?? row.weeklyPlan),
      weeklyActual: normalizeNumber(row["周毛利实际RMB"] ?? row.weeklyActual),
      completion:
        row["完成率"] === null || row["完成率"] === undefined
          ? null
          : normalizeNumber(row["完成率"]),
      gap: normalizeNumber(row["差异"] ?? row.gap),
      weekNo: normalizeInteger(row["周序号"] ?? row.weekNo),
    }))
    .filter((row) => row.country && row.asin);
}

function normalizeMonthlyTrackRows(rows) {
  return rows
    .map((row) => ({
      month: normalizeText(row["月份"] ?? row.month),
      country: normalizeText(row["国家"] ?? row.country),
      asin: normalizeText(row["ASIN"] ?? row["asin"] ?? row.asin),
      product: normalizeText(row["品名"] ?? row.product),
      marketLevel: normalizeText(row["市场等级"] ?? row.marketLevel),
      listingLevel: normalizeText(row["链接分级"] ?? row.listingLevel),
      owner: normalizeText(row["负责人"] ?? row.owner ?? row["文本 5"]),
      monthlyPlan: normalizeNumber(row["月毛利预期RMB"] ?? row.monthlyPlan),
      monthlyActual: normalizeNumber(row["月毛利实际RMB"] ?? row.monthlyActual),
      completion:
        row["完成率"] === null || row["完成率"] === undefined
          ? null
          : normalizeNumber(row["完成率"]),
      gap: normalizeNumber(row["差异"] ?? row.gap),
    }))
    .filter((row) => row.country && row.asin);
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/,/g, "").trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeInteger(value) {
  const n = normalizeNumber(value);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object") {
    if (value.text !== undefined) return String(value.text).trim();
    if (value.name !== undefined) return String(value.name).trim();
    if (value.value !== undefined) return String(value.value).trim();
  }
  return String(value).trim();
}

function normalizeDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number" && value > 20000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  if (!text) return "";
  const simpleMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (simpleMatch) {
    return `${simpleMatch[1]}-${simpleMatch[2].padStart(2, "0")}-${simpleMatch[3].padStart(2, "0")}`;
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function parseWeekRange(weekRange) {
  const match = weekRange?.match(
    /(\d{4}-\d{2}-\d{2})\s*[~～-]\s*(\d{4}-\d{2}-\d{2})/
  );
  return match
    ? { start: match[1], end: match[2] }
    : { start: "", end: "" };
}

function renderDashboard() {
  if (!state.dataset) {
    renderEmptyState();
    return;
  }

  const model = buildDashboardModel(state.dataset, state.manualDate);
  validateDashboardModel(model);
  renderHero(model);
  renderKpis(model);
  renderCountryBars(model.countrySummary);
  renderCountryTable(model.countrySummary);
  hydrateFilters(model);
  renderWarnings(model.filteredWarningRows);
  renderDetailTable(model.filteredDetailRows, model.detailSummary);
  renderWeeklyTable(model.filteredWeeklyRows, model.weeklySummary);
  renderMonthlyTable(model.filteredMonthlyRows, model.monthlySummary);
}

function buildDashboardModel(dataset, manualDate) {
  const availableDates = [...new Set(dataset.profitRows.map((row) => row.date))]
    .filter(Boolean)
    .sort();
  if (!availableDates.length) {
    throw new Error("利润表中没有可用日期");
  }

  const latestDate = availableDates[availableDates.length - 1];
  const suggestedDate =
    availableDates.length > 1 ? availableDates[availableDates.length - 2] : latestDate;
  const effectiveDate =
    manualDate && availableDates.includes(manualDate) ? manualDate : suggestedDate;
  const compareDate =
    [...availableDates].reverse().find((date) => date < effectiveDate) || effectiveDate;

  const currentWeek =
    dataset.planRows.find(
      (row) =>
        row.weekStart &&
        row.weekEnd &&
        row.weekStart <= effectiveDate &&
        row.weekEnd >= effectiveDate
    ) ||
    dataset.planRows.find((row) => row.weekEnd >= effectiveDate) ||
    dataset.planRows[dataset.planRows.length - 1];

  const currentWeekNo = currentWeek?.weekNo || 0;
  const currentWeekRange = currentWeek?.weekRange || "";
  const currentWeekRows = dataset.planRows.filter((row) => row.weekNo === currentWeekNo);
  const ownerMap = new Map();
  (dataset.ownerRows || []).forEach((row) => {
    const key = `${row.country}__${row.asin}__${row.product}`;
    if (!ownerMap.has(key) && row.owner) ownerMap.set(key, row.owner);
    const looseKey = `${row.country}__${row.asin}`;
    if (!ownerMap.has(looseKey) && row.owner) ownerMap.set(looseKey, row.owner);
  });

  const shoeDryerProfitRows = dataset.profitRows.filter((row) =>
    (row.product || "").includes("烘鞋器")
  );

  const countries = [
    ...new Set(
      [...shoeDryerProfitRows, ...currentWeekRows].map((row) => row.country).filter(Boolean)
    ),
  ];

  const sumProfit = (criteria) =>
    round(
      shoeDryerProfitRows
        .filter(criteria)
        .reduce((sum, row) => sum + (row.grossProfitRmb || 0), 0)
    );

  const groupedPlanRows = new Map();
  currentWeekRows.forEach((row) => {
    const key = `${row.country}__${row.asin}`;
    if (!groupedPlanRows.has(key)) {
      groupedPlanRows.set(key, {
        country: row.country,
        asin: row.asin,
        product: row.product,
        marketLevel: row.marketLevel || "",
        listingLevel: row.listingLevel || "",
        owner:
          ownerMap.get(`${row.country}__${row.asin}__${row.product}`) ||
          ownerMap.get(`${row.country}__${row.asin}`) ||
          "",
        weeklyPlan: 0,
        dailyPlan: 0,
      });
    }
    const target = groupedPlanRows.get(key);
    target.weeklyPlan += row.weeklyProfitRmb;
    target.dailyPlan += row.dailyProfitRmb;
    if (!target.product && row.product) target.product = row.product;
  });

  const detailRows = [...groupedPlanRows.values()]
    .map((row) => {
      const latest = sumProfit(
        (item) =>
          item.date === effectiveDate &&
          item.country === row.country &&
          item.asin === row.asin
      );
      const previous = sumProfit(
        (item) =>
          item.date === compareDate &&
          item.country === row.country &&
          item.asin === row.asin
      );
      const diff = round(latest - previous);
      const diffPct = calculateTrendRate(diff, previous);
      const vsPlan = round(latest - row.dailyPlan);
      return {
        ...row,
        latest,
        previous,
        diff,
        diffPct,
        dailyPlan: round(row.dailyPlan),
        vsPlan,
        status: `${diff > 0 ? "上涨" : diff < 0 ? "下跌" : "持平"} | ${
          vsPlan >= 0 ? "超预期" : "低于预期"
        }`,
      };
    })
    .sort((a, b) => a.country.localeCompare(b.country) || b.latest - a.latest);

  const weeklyTrackSource = (dataset.weeklyTrackRows || []).filter(
    (row) => row.weekNo === currentWeekNo
  );
  const weeklyRows = (
    weeklyTrackSource.length
      ? weeklyTrackSource.map((row) => ({
          country: row.country,
          asin: row.asin,
          product: row.product,
          marketLevel: row.marketLevel || "",
          listingLevel: row.listingLevel || "",
          owner:
            row.owner ||
            ownerMap.get(`${row.country}__${row.asin}__${row.product}`) ||
            ownerMap.get(`${row.country}__${row.asin}`) ||
            "",
          weeklyPlan: round(row.weeklyPlan),
          weeklyActual: round(row.weeklyActual),
          completion: row.completion,
          gap: round(row.gap),
        }))
      : [...groupedPlanRows.values()].map((row) => {
          const weeklyActual = sumProfit(
            (item) =>
              item.country === row.country &&
              item.asin === row.asin &&
              item.date >= currentWeek?.weekStart &&
              item.date <= currentWeek?.weekEnd
          );
          const completion = row.weeklyPlan === 0 ? 0 : weeklyActual / row.weeklyPlan;
          return {
            country: row.country,
            asin: row.asin,
            product: row.product,
            marketLevel: row.marketLevel || "",
            listingLevel: row.listingLevel || "",
            owner: row.owner || "",
            weeklyPlan: round(row.weeklyPlan),
            weeklyActual,
            completion,
            gap: round(weeklyActual - row.weeklyPlan),
          };
        })
  ).sort((a, b) => a.country.localeCompare(b.country) || (b.completion || 0) - (a.completion || 0));

  const weeklyCountryMap = new Map();
  weeklyRows.forEach((row) => {
    if (!weeklyCountryMap.has(row.country)) {
      weeklyCountryMap.set(row.country, {
        weeklyPlan: 0,
        weeklyActual: 0,
      });
    }
    const target = weeklyCountryMap.get(row.country);
    target.weeklyPlan += row.weeklyPlan;
    target.weeklyActual += row.weeklyActual;
  });

  const countrySummary = countries
    .map((country) => {
      const latest = sumProfit(
        (row) => row.date === effectiveDate && row.country === country
      );
      const previous = sumProfit(
        (row) => row.date === compareDate && row.country === country
      );
      const diff = round(latest - previous);
      const diffPct = calculateTrendRate(diff, previous);
      const weeklyCountry = weeklyCountryMap.get(country) || {
        weeklyPlan: 0,
        weeklyActual: 0,
      };
      const weeklyPlan = round(weeklyCountry.weeklyPlan);
      const weeklyActual = round(weeklyCountry.weeklyActual);
      const completion = weeklyPlan === 0 ? 0 : weeklyActual / weeklyPlan;
      return {
        country,
        latest,
        previous,
        diff,
        diffPct,
        weeklyPlan,
        weeklyActual,
        completion,
      };
    })
    .sort((a, b) => (b.latest - a.latest) || a.country.localeCompare(b.country));
  const sortedCountrySummary = sortRows(countrySummary, state.countrySort);

  const filteredDetailRows = applyTableFilters(detailRows, state.detailFilters);
  const sortedDetailRows = sortRows(filteredDetailRows, state.detailSort);
  const detailSummary = buildDetailSummary(sortedDetailRows);

  const filteredWeeklyRows = applyTableFilters(weeklyRows, state.weeklyFilters);
  const sortedWeeklyRows = sortRows(filteredWeeklyRows, state.weeklySort);
  const weeklySummary = buildWeeklySummary(sortedWeeklyRows);
  const overallWeeklySummary = buildWeeklySummary(weeklyRows);
  const monthlyRows = (dataset.monthlyTrackRows || []).map((row) => ({
    month: row.month,
    country: row.country,
    asin: row.asin,
    product: row.product,
    marketLevel: row.marketLevel || "",
    listingLevel: row.listingLevel || "",
    owner:
      row.owner ||
      ownerMap.get(`${row.country}__${row.asin}__${row.product}`) ||
      ownerMap.get(`${row.country}__${row.asin}`) ||
      "",
    monthlyPlan: round(row.monthlyPlan),
    monthlyActual: round(row.monthlyActual),
    completion: row.completion,
    gap: round(row.gap),
  }));
  const filteredMonthlyRows = applyTableFilters(monthlyRows, state.monthlyFilters);
  const sortedMonthlyRows = sortRows(filteredMonthlyRows, state.monthlySort);
  const monthlySummary = buildMonthlySummary(sortedMonthlyRows);
  const currentMonthLabel = `${Number(effectiveDate.slice(5, 7))}月`;
  const currentMonthRows = monthlyRows.filter((row) => row.month === currentMonthLabel);
  const currentMonthSummary = buildMonthlySummary(
    currentMonthRows.length ? currentMonthRows : monthlyRows
  );

  const latestTotal = round(countrySummary.reduce((sum, row) => sum + row.latest, 0));
  const previousTotal = round(countrySummary.reduce((sum, row) => sum + row.previous, 0));
  const totalDiff = round(latestTotal - previousTotal);
  const totalDiffPct = calculateTrendRate(totalDiff, previousTotal);
  const totalWeekPlan = overallWeeklySummary.weeklyPlan;
  const totalWeekActual = overallWeeklySummary.weeklyActual;
  const totalWeekCompletion = overallWeeklySummary.completion;

  const warningRows = buildWarningRows(detailRows, weeklyRows, shoeDryerProfitRows, effectiveDate);
  const filteredWarningRows = warningRows.filter((row) => {
    if (state.warningFilters.risk && row.riskLevelLabel !== state.warningFilters.risk) return false;
    if (state.warningFilters.market && row.marketLevel !== state.warningFilters.market) return false;
    if (state.warningFilters.country && row.country !== state.warningFilters.country) return false;
    if (state.warningFilters.asin && row.asin !== state.warningFilters.asin) return false;
    if (state.warningFilters.product && row.product !== state.warningFilters.product) return false;
    if (state.warningFilters.owner && (row.owner || "") !== state.warningFilters.owner) return false;
    return true;
  });
  const sortedWarningRows = sortRows(filteredWarningRows, state.warningSort);

  return {
    sourceMeta: dataset.sourceMeta,
    effectiveDate,
    compareDate,
    latestTotal,
    previousTotal,
    totalDiff,
    totalDiffPct,
    currentWeekNo,
    currentWeekRange,
    totalWeekPlan,
    totalWeekActual,
    totalWeekCompletion,
    countrySummary: sortedCountrySummary,
    detailRows,
    weeklyRows,
    filteredDetailRows: sortedDetailRows,
    filteredWeeklyRows: sortedWeeklyRows,
    detailSummary,
    weeklySummary,
    overallWeeklySummary,
    monthlyRows,
    filteredMonthlyRows: sortedMonthlyRows,
    monthlySummary,
    currentMonthLabel,
    currentMonthSummary,
    warningRows,
    filteredWarningRows: sortedWarningRows,
  };
}

function renderEmptyState() {
  refs.heroMeta.textContent = "等待数据加载";
  refs.heroPill.textContent = "未加载";
}

function renderHero(model) {
  refs.heroMeta.textContent = `最新日期 ${model.effectiveDate} ｜ 对比日期 ${model.compareDate} ｜ 当前周 ${model.currentWeekRange || "-"}`;
  refs.heroPill.textContent =
    model.sourceMeta.source === "feishu" ? "Feishu Connected" : "Excel Upload";
}

function renderKpis(model) {
  const cards = [
    {
      value: model.effectiveDate,
      subvalue: formatCurrency(model.latestTotal),
      tone: model.latestTotal,
    },
    {
      value: model.compareDate,
      subvalue: formatCurrency(model.previousTotal ?? 0),
      tone: model.previousTotal ?? 0,
    },
    {
      value: formatCurrency(model.totalDiff),
      subvalue: formatPercent(model.totalDiffPct),
      tone: model.totalDiff,
    },
    {
      value: model.currentWeekRange || "-",
      subvalue: `W${model.currentWeekNo || "-"}`,
      tone: null,
    },
    {
      value: formatPercent(model.totalWeekCompletion),
      subvalue: `${formatCurrency(model.totalWeekActual)} / ${formatCurrency(model.totalWeekPlan)}`,
      tone: model.totalWeekCompletion,
    },
    {
      value: formatPercent(model.currentMonthSummary?.completion ?? 0),
      subvalue: `${formatCurrency(model.currentMonthSummary?.monthlyActual ?? 0)} / ${formatCurrency(model.currentMonthSummary?.monthlyPlan ?? 0)}`,
      tone: model.currentMonthSummary?.completion ?? 0,
    },
  ];
  const values = refs.kpiGrid.querySelectorAll(".kpi-value");
  const subvalues = refs.kpiGrid.querySelectorAll(".kpi-subvalue");
  values.forEach((node, index) => {
    const card = cards[index];
    node.textContent = card?.value || "-";
    node.className = `kpi-value ${toneClass(card?.tone)}`;
  });
  subvalues.forEach((node, index) => {
    const card = cards[index];
    node.textContent = card?.subvalue || "-";
    node.className = `kpi-subvalue ${toneClass(card?.tone)}`;
  });
}

function renderCountryBars(rows) {
  if (!rows.length) {
    refs.countryBars.innerHTML = "没有可展示的站点汇总";
    refs.countryBars.classList.add("empty-state");
    return;
  }
  refs.countryBars.classList.remove("empty-state");
  const maxAbs = Math.max(...rows.map((row) => Math.abs(row.latest)), 1);
  refs.countryBars.innerHTML = rows
    .map((row) => {
      const width = Math.max((Math.abs(row.latest) / maxAbs) * 100, row.latest !== 0 ? 4 : 0);
      return `
        <div class="country-bar">
          <div class="country-bar__meta">
            <strong>${row.country}</strong>
            <span>${formatCurrency(row.latest)} ｜ 周完成 ${formatPercent(row.completion)}</span>
          </div>
          <div class="country-bar__track">
            <div class="country-bar__fill ${row.latest < 0 ? "is-negative" : ""}" style="width:${width}%"></div>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderCountryTable(rows) {
  refs.summaryCount.textContent = `${rows.length} 个站点`;
  if (!rows.length) {
    refs.countryTableBody.innerHTML =
      '<tr><td colspan="8" class="table-placeholder">没有站点数据</td></tr>';
    return;
  }

  refs.countryTableBody.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.country}</td>
        <td class="${toneClass(row.latest)}">${formatCurrency(row.latest)}</td>
        <td class="${toneClass(row.previous)}">${formatCurrency(row.previous)}</td>
        <td class="${toneClass(row.diff)}">${formatCurrency(row.diff)}</td>
        <td class="${toneClass(row.diffPct)}">${formatPercent(row.diffPct)}</td>
        <td class="${toneClass(row.weeklyPlan)}">${formatCurrency(row.weeklyPlan)}</td>
        <td class="${toneClass(row.weeklyActual)}">${formatCurrency(row.weeklyActual)}</td>
        <td class="${toneClass(row.completion)}">${formatPercent(row.completion)}</td>
      </tr>
    `
    )
    .join("");
}

function renderWarnings(rows) {
  refs.warningCount.textContent = `${rows.length} 条`;
  refs.warningTableCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.warningTopCards.innerHTML = "当前没有需要重点处理的预警 SKU";
    refs.warningTopCards.classList.add("empty-state");
    refs.warningTableBody.innerHTML =
      '<tr><td colspan="15" class="table-placeholder">当前没有需要重点处理的预警 SKU</td></tr>';
    return;
  }

  refs.warningTopCards.classList.remove("empty-state");
  refs.warningTopCards.innerHTML = rows
    .slice(0, 5)
    .map(
      (row) => `
      <article class="highlight-card ${row.riskLevelClass}">
        <h4>${row.country} | ${row.asin}</h4>
        <strong>${row.product}</strong>
        <span>${row.owner || "未分配"} ｜ ${row.riskLevelLabel}</span>
        <span>${formatCurrency(row.latest)} ｜ 环比 ${formatCurrency(row.diff)} ｜ vs预期 ${formatCurrency(row.vsPlan)}</span>
        <span>${row.aiSummary}</span>
      </article>
    `
    )
    .join("");

  refs.warningTableBody.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${renderRiskPill(row.riskLevelLabel, row.riskLevelClass)}</td>
        <td>${row.country}</td>
        <td>${row.asin}</td>
        <td>${row.product || "-"}</td>
        <td>${row.marketLevel || "-"}</td>
        <td>${row.listingLevel || "-"}</td>
        <td>${row.owner || "-"}</td>
        <td class="${toneClass(row.latest)}">${formatCurrency(row.latest)}</td>
        <td class="${toneClass(row.diff)}">${formatCurrency(row.diff)}</td>
        <td class="${toneClass(row.vsPlan)}">${formatCurrency(row.vsPlan)}</td>
        <td class="${toneClass(row.weeklyCompletion)}">${formatPercent(row.weeklyCompletion)}</td>
        <td>${row.issueTypes.join(" / ")}</td>
        <td>${row.aiAdvice}</td>
        <td>${row.priority}</td>
        <td>${row.followOwner}</td>
      </tr>
    `
    )
    .join("");
}

function renderDetailTable(rows, summary) {
  refs.detailCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.detailTableBody.innerHTML =
      '<tr><td colspan="13" class="table-placeholder">没有 ASIN 明细</td></tr>';
    return;
  }

  refs.detailTableBody.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.country}</td>
        <td>${row.asin}</td>
        <td>${row.product || "-"}</td>
        <td>${row.marketLevel || "-"}</td>
        <td>${row.listingLevel || "-"}</td>
        <td>${row.owner || "-"}</td>
        <td class="${toneClass(row.latest)}">${formatCurrency(row.latest)}</td>
        <td class="${toneClass(row.previous)}">${formatCurrency(row.previous)}</td>
        <td class="${toneClass(row.diff)}">${formatCurrency(row.diff)}</td>
        <td class="${toneClass(row.diffPct)}">${formatPercent(row.diffPct)}</td>
        <td class="${toneClass(row.dailyPlan)}">${formatCurrency(row.dailyPlan)}</td>
        <td class="${toneClass(row.vsPlan)}">${formatCurrency(row.vsPlan)}</td>
        <td>${row.status}</td>
      </tr>
    `
    )
    .join("") + renderDetailSummaryRow(summary);
}

function renderWeeklyTable(rows, summary) {
  refs.weeklyCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.weeklyTableBody.innerHTML =
      '<tr><td colspan="10" class="table-placeholder">没有周进度数据</td></tr>';
    return;
  }

  refs.weeklyTableBody.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.country}</td>
        <td>${row.asin}</td>
        <td>${row.product || "-"}</td>
        <td>${row.marketLevel || "-"}</td>
        <td>${row.listingLevel || "-"}</td>
        <td>${row.owner || "-"}</td>
        <td class="${toneClass(row.weeklyPlan)}">${formatCurrency(row.weeklyPlan)}</td>
        <td class="${toneClass(row.weeklyActual)}">${formatCurrency(row.weeklyActual)}</td>
        <td class="${toneClass(row.completion)}">${formatPercent(row.completion)}</td>
        <td class="${toneClass(row.gap)}">${formatCurrency(row.gap)}</td>
      </tr>
    `
    )
    .join("") + renderWeeklySummaryRow(summary);
}

function renderMonthlyTable(rows, summary) {
  refs.monthlyCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.monthlyTableBody.innerHTML =
      '<tr><td colspan="11" class="table-placeholder">没有月进度数据</td></tr>';
    return;
  }

  refs.monthlyTableBody.innerHTML =
    rows
      .map(
        (row) => `
      <tr>
        <td>${row.month || "-"}</td>
        <td>${row.country}</td>
        <td>${row.asin}</td>
        <td>${row.product || "-"}</td>
        <td>${row.marketLevel || "-"}</td>
        <td>${row.listingLevel || "-"}</td>
        <td>${row.owner || "-"}</td>
        <td class="${toneClass(row.monthlyPlan)}">${formatCurrency(row.monthlyPlan)}</td>
        <td class="${toneClass(row.monthlyActual)}">${formatCurrency(row.monthlyActual)}</td>
        <td class="${toneClass(row.completion)}">${formatPercent(row.completion)}</td>
        <td class="${toneClass(row.gap)}">${formatCurrency(row.gap)}</td>
      </tr>
    `
      )
      .join("") + renderMonthlySummaryRow(summary);
}

function renderProgressChip(status) {
  const className =
    status === "达成/超额" ? "done" : status === "接近达成" ? "warn" : "todo";
  return `<span class="progress-chip ${className}">${status}</span>`;
}

function renderRiskPill(label, className) {
  const levelClass =
    className === "warn-high" ? "high" : className === "warn-mid" ? "medium" : "low";
  return `<span class="risk-pill ${levelClass}">${label}</span>`;
}

function applyTableFilters(rows, filters) {
  return rows.filter((row) => {
    if (filters.month && row.month !== filters.month) return false;
    if (filters.market && row.marketLevel !== filters.market) return false;
    if (filters.country && row.country !== filters.country) return false;
    if (filters.asin && row.asin !== filters.asin) return false;
    if (filters.product && row.product !== filters.product) return false;
    if (filters.owner && (row.owner || "") !== filters.owner) return false;
    return true;
  });
}

function sortRows(rows, sortConfig) {
  const sorted = [...rows];
  const direction = sortConfig.direction === "asc" ? 1 : -1;
  sorted.sort((a, b) => {
    const av = a[sortConfig.key];
    const bv = b[sortConfig.key];
    if (typeof av === "number" && typeof bv === "number") {
      return (av - bv) * direction;
    }
    return String(av || "").localeCompare(String(bv || "")) * direction;
  });
  return sorted;
}

function buildDetailSummary(rows) {
  const latest = round(rows.reduce((sum, row) => sum + (row.latest || 0), 0));
  const previous = round(rows.reduce((sum, row) => sum + (row.previous || 0), 0));
  const diff = round(rows.reduce((sum, row) => sum + (row.diff || 0), 0));
  const dailyPlan = round(rows.reduce((sum, row) => sum + (row.dailyPlan || 0), 0));
  const vsPlan = round(rows.reduce((sum, row) => sum + (row.vsPlan || 0), 0));
  return {
    latest,
    previous,
    diff,
    diffPct: calculateTrendRate(diff, previous),
    dailyPlan,
    vsPlan,
  };
}

function buildWeeklySummary(rows) {
  const weeklyPlan = round(rows.reduce((sum, row) => sum + (row.weeklyPlan || 0), 0));
  const weeklyActual = round(rows.reduce((sum, row) => sum + (row.weeklyActual || 0), 0));
  const gap = round(rows.reduce((sum, row) => sum + (row.gap || 0), 0));
  return {
    weeklyPlan,
    weeklyActual,
    completion: weeklyPlan === 0 ? 0 : weeklyActual / weeklyPlan,
    gap,
  };
}

function buildWarningRows(detailRows, weeklyRows, profitRows, effectiveDate) {
  const weeklyMap = new Map(
    weeklyRows.map((row) => [`${row.country}__${row.asin}`, row])
  );

  const seriesMap = new Map();
  profitRows.forEach((row) => {
    const key = `${row.country}__${row.asin}`;
    if (!seriesMap.has(key)) seriesMap.set(key, []);
    seriesMap.get(key).push({ date: row.date, value: row.grossProfitRmb || 0 });
  });
  seriesMap.forEach((items) => items.sort((a, b) => a.date.localeCompare(b.date)));

  return detailRows
    .map((row) => {
      const weekly = weeklyMap.get(`${row.country}__${row.asin}`) || {};
      const recentSeries = (seriesMap.get(`${row.country}__${row.asin}`) || [])
        .filter((item) => item.date <= effectiveDate)
        .slice(-3);
      const issueTypes = [];
      let riskScore = 0;

      if (row.diff < 0) {
        issueTypes.push("环比下滑");
        riskScore += Math.min(Math.abs(row.diff) / 100, 3);
      }
      if (row.vsPlan < 0) {
        issueTypes.push("低于预期");
        riskScore += Math.min(Math.abs(row.vsPlan) / 100, 3);
      }
      if (row.latest < 0) {
        issueTypes.push("负毛利");
        riskScore += 2.5;
      }
      if (
        recentSeries.length === 3 &&
        recentSeries[0].value > recentSeries[1].value &&
        recentSeries[1].value > recentSeries[2].value
      ) {
        issueTypes.push("连续下滑");
        riskScore += 2;
      }
      if ((weekly.completion || 0) < 0.6 && (weekly.weeklyPlan || 0) > 500) {
        issueTypes.push("高预期低达成");
        riskScore += 2.5;
      }
      if (Math.abs(row.diff) > 300 && Math.abs(row.latest) < 50) {
        issueTypes.push("异常波动");
        riskScore += 1.5;
      }

      if (!issueTypes.length) return null;

      const riskLevelLabel =
        riskScore >= 6 ? "高风险" : riskScore >= 3 ? "中风险" : "低风险";
      const riskLevelClass =
        riskScore >= 6 ? "warn-high" : riskScore >= 3 ? "warn-mid" : "warn-low";
      const priority = riskScore >= 6 ? "今日处理" : riskScore >= 3 ? "本周优先" : "持续观察";
      const aiSummary = buildAiSummary(row, weekly, issueTypes);
      const aiAdvice = buildAiAdvice(row, weekly, issueTypes);

      return {
        ...row,
        marketLevel: row.marketLevel || "",
        listingLevel: row.listingLevel || "",
        weeklyCompletion: weekly.completion || 0,
        issueTypes,
        riskScore,
        riskLevelLabel,
        riskLevelClass,
        aiSummary,
        aiAdvice,
        priority,
        followOwner: row.owner || "待分配",
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.riskScore - a.riskScore || a.country.localeCompare(b.country));
}

function buildAiSummary(row, weekly, issueTypes) {
  if (issueTypes.includes("负毛利") && issueTypes.includes("低于预期")) {
    return "今日毛利转弱且未达预期，属于优先处理异常。";
  }
  if (issueTypes.includes("高预期低达成")) {
    return "目标高但周进度偏低，建议优先检查投放和转化效率。";
  }
  if (issueTypes.includes("连续下滑")) {
    return "最近 3 天连续走弱，建议尽快排查价格、广告和库存。";
  }
  return `当前主要问题为${issueTypes.join("、")}，需要尽快跟进。`;
}

function buildAiAdvice(row, weekly, issueTypes) {
  const actions = [];
  if (issueTypes.includes("负毛利")) {
    actions.push("优先检查广告花费、折扣和定价是否异常");
  }
  if (issueTypes.includes("低于预期")) {
    actions.push("对比单日预期与实际流量转化，判断是否需要补量");
  }
  if (issueTypes.includes("连续下滑")) {
    actions.push("连续 2-3 天观察 Sessions、CVR 和广告占比");
  }
  if (issueTypes.includes("高预期低达成")) {
    actions.push(`本周完成率仅 ${formatPercent(weekly.completion || 0)}，优先调整投放和节奏`);
  }
  if (!actions.length) {
    actions.push("建议继续观察未来 1-2 天毛利与转化变化");
  }
  return actions.join("；");
}

function buildMonthlySummary(rows) {
  const monthlyPlan = round(rows.reduce((sum, row) => sum + (row.monthlyPlan || 0), 0));
  const monthlyActual = round(rows.reduce((sum, row) => sum + (row.monthlyActual || 0), 0));
  const gap = round(rows.reduce((sum, row) => sum + (row.gap || 0), 0));
  return {
    monthlyPlan,
    monthlyActual,
    completion: monthlyPlan === 0 ? 0 : monthlyActual / monthlyPlan,
    gap,
  };
}

function renderDetailSummaryRow(summary) {
  if (!summary) return "";
  return `
    <tr class="summary-row">
      <td>汇总</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td class="${toneClass(summary.latest)}">${formatCurrency(summary.latest)}</td>
      <td class="${toneClass(summary.previous)}">${formatCurrency(summary.previous)}</td>
      <td class="${toneClass(summary.diff)}">${formatCurrency(summary.diff)}</td>
      <td class="${toneClass(summary.diffPct)}">${formatPercent(summary.diffPct)}</td>
      <td class="${toneClass(summary.dailyPlan)}">${formatCurrency(summary.dailyPlan)}</td>
      <td class="${toneClass(summary.vsPlan)}">${formatCurrency(summary.vsPlan)}</td>
      <td>-</td>
    </tr>
  `;
}

function renderWeeklySummaryRow(summary) {
  if (!summary) return "";
  return `
    <tr class="summary-row">
      <td>汇总</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td class="${toneClass(summary.weeklyPlan)}">${formatCurrency(summary.weeklyPlan)}</td>
      <td class="${toneClass(summary.weeklyActual)}">${formatCurrency(summary.weeklyActual)}</td>
      <td class="${toneClass(summary.completion)}">${formatPercent(summary.completion)}</td>
      <td class="${toneClass(summary.gap)}">${formatCurrency(summary.gap)}</td>
    </tr>
  `;
}

function renderMonthlySummaryRow(summary) {
  if (!summary) return "";
  return `
    <tr class="summary-row">
      <td>汇总</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td class="${toneClass(summary.monthlyPlan)}">${formatCurrency(summary.monthlyPlan)}</td>
      <td class="${toneClass(summary.monthlyActual)}">${formatCurrency(summary.monthlyActual)}</td>
      <td class="${toneClass(summary.completion)}">${formatPercent(summary.completion)}</td>
      <td class="${toneClass(summary.gap)}">${formatCurrency(summary.gap)}</td>
    </tr>
  `;
}

function hydrateFilters(model) {
  populateFilter(refs.warningRiskFilter, uniqueValues(model.warningRows, "riskLevelLabel"), "全部风险等级", state.warningFilters.risk);
  populateFilter(refs.warningMarketFilter, uniqueValues(model.warningRows, "marketLevel"), "全部市场等级", state.warningFilters.market);
  populateFilter(refs.warningCountryFilter, uniqueValues(model.warningRows, "country"), "全部国家", state.warningFilters.country);
  populateFilter(refs.warningAsinFilter, uniqueValues(model.warningRows, "asin"), "全部 ASIN", state.warningFilters.asin);
  populateFilter(refs.warningProductFilter, uniqueValues(model.warningRows, "product"), "全部品名", state.warningFilters.product);
  populateFilter(refs.warningOwnerFilter, uniqueValues(model.warningRows, "owner"), "全部负责人", state.warningFilters.owner);

  populateFilter(refs.detailMarketFilter, uniqueValues(model.detailRows, "marketLevel"), "全部市场等级", state.detailFilters.market);
  populateFilter(refs.detailCountryFilter, uniqueValues(model.detailRows, "country"), "全部国家", state.detailFilters.country);
  populateFilter(refs.detailAsinFilter, uniqueValues(model.detailRows, "asin"), "全部 ASIN", state.detailFilters.asin);
  populateFilter(refs.detailProductFilter, uniqueValues(model.detailRows, "product"), "全部品名", state.detailFilters.product);
  populateFilter(refs.detailOwnerFilter, uniqueValues(model.detailRows, "owner"), "全部负责人", state.detailFilters.owner);

  populateFilter(refs.weeklyMarketFilter, uniqueValues(model.weeklyRows, "marketLevel"), "全部市场等级", state.weeklyFilters.market);
  populateFilter(refs.weeklyCountryFilter, uniqueValues(model.weeklyRows, "country"), "全部国家", state.weeklyFilters.country);
  populateFilter(refs.weeklyAsinFilter, uniqueValues(model.weeklyRows, "asin"), "全部 ASIN", state.weeklyFilters.asin);
  populateFilter(refs.weeklyProductFilter, uniqueValues(model.weeklyRows, "product"), "全部品名", state.weeklyFilters.product);
  populateFilter(refs.weeklyOwnerFilter, uniqueValues(model.weeklyRows, "owner"), "全部负责人", state.weeklyFilters.owner);

  populateFilter(refs.monthlyMonthFilter, uniqueValues(model.monthlyRows, "month"), "全部月份", state.monthlyFilters.month);
  populateFilter(refs.monthlyMarketFilter, uniqueValues(model.monthlyRows, "marketLevel"), "全部市场等级", state.monthlyFilters.market);
  populateFilter(refs.monthlyCountryFilter, uniqueValues(model.monthlyRows, "country"), "全部国家", state.monthlyFilters.country);
  populateFilter(refs.monthlyAsinFilter, uniqueValues(model.monthlyRows, "asin"), "全部 ASIN", state.monthlyFilters.asin);
  populateFilter(refs.monthlyOwnerFilter, uniqueValues(model.monthlyRows, "owner"), "全部负责人", state.monthlyFilters.owner);
}

function populateFilter(select, values, defaultLabel, selectedValue) {
  const current = selectedValue || "";
  const options = [`<option value="">${defaultLabel}</option>`]
    .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
    .join("");
  select.innerHTML = options;
  select.value = values.includes(current) ? current : "";
}

function uniqueValues(rows, key) {
  return [...new Set(rows.map((row) => row[key]).filter(Boolean))].sort((a, b) =>
    String(a).localeCompare(String(b))
  );
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function toneClass(value) {
  if (value === null || value === undefined || value === "") return "value-neutral";
  if (value > 0) return "value-positive";
  if (value < 0) return "value-negative";
  return "value-neutral";
}

function formatCurrency(value) {
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value || 0);
}

function formatPercent(value) {
  if (value === null || value === undefined || value === "") return "-";
  return `${(value * 100).toFixed(1)}%`;
}

function calculateTrendRate(diff, previous) {
  if (previous === 0 || previous === null || previous === undefined) return null;
  return diff / Math.abs(previous);
}

function round(value) {
  return Number((value || 0).toFixed(2));
}

function validateDashboardModel(model) {
  const issues = [];

  const inspectRows = (rows, scope) => {
    rows.forEach((row) => {
      if (row.diff !== undefined && row.previous !== undefined) {
        const expectedDiff = round((row.latest || 0) - (row.previous || 0));
        if (expectedDiff !== round(row.diff || 0)) {
          issues.push(
            `${scope} ${row.country || "-"} ${row.asin || ""} 差值校验失败：latest - previous != diff`
          );
        }
      }

      if (row.diffPct !== null && row.diffPct !== undefined) {
        if (row.diff > 0 && row.diffPct < 0) {
          issues.push(
            `${scope} ${row.country || "-"} ${row.asin || ""} 百分比方向与上涨不一致`
          );
        }
        if (row.diff < 0 && row.diffPct > 0) {
          issues.push(
            `${scope} ${row.country || "-"} ${row.asin || ""} 百分比方向与下跌不一致`
          );
        }
      }

      if (row.status) {
        if (row.diff > 0 && !row.status.startsWith("上涨")) {
          issues.push(`${scope} ${row.country || "-"} ${row.asin || ""} 状态应为上涨`);
        }
        if (row.diff < 0 && !row.status.startsWith("下跌")) {
          issues.push(`${scope} ${row.country || "-"} ${row.asin || ""} 状态应为下跌`);
        }
        if (row.diff === 0 && !row.status.startsWith("持平")) {
          issues.push(`${scope} ${row.country || "-"} ${row.asin || ""} 状态应为持平`);
        }
      }
    });
  };

  inspectRows(model.countrySummary, "站点总览");
  inspectRows(model.detailRows, "ASIN明细");

  model.weeklyRows.forEach((row) => {
    const expectedGap = round((row.weeklyActual || 0) - (row.weeklyPlan || 0));
    if (expectedGap !== round(row.gap || 0)) {
      issues.push(`${row.country || "-"} ${row.asin || ""} 周进度差异校验失败`);
    }
  });

  if (
    round(model.totalWeekCompletion || 0) !==
    round(model.overallWeeklySummary?.completion || 0)
  ) {
    issues.push("顶部当前周完成率 与 Q2周毛利进度汇总完成率 不一致");
  }

  if (issues.length) {
    console.warn("Dashboard validation issues:", issues);
    showToast(`检测到 ${issues.length} 条数据校验提醒，请打开控制台查看详情`);
  }
}

function hydrateDateInput() {
  if (!state.dataset) return;
  const availableDates = [...new Set(state.dataset.profitRows.map((row) => row.date))]
    .filter(Boolean)
    .sort();
  refs.manualDate.min = availableDates[0] || "";
  refs.manualDate.max = availableDates[availableDates.length - 1] || "";
  if (!state.manualDate && availableDates.length > 1) {
    refs.manualDate.value = availableDates[availableDates.length - 2];
  }
}

function setStatus(label, className) {
  refs.sourceStatus.textContent = label;
  refs.sourceStatus.className = `badge ${className}`;
}

let toastTimer = null;
function showToast(message) {
  refs.toast.textContent = message;
  refs.toast.classList.remove("hidden");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => refs.toast.classList.add("hidden"), 3200);
}

if (IS_STANDALONE) {
  const feishuOption = document.querySelector('input[name="sourceMode"][value="feishu"]');
  if (feishuOption) {
    feishuOption.disabled = true;
    feishuOption.closest(".mode-option")?.setAttribute("title", "离线单文件版仅支持 Excel 上传");
  }
  refs.feishuPanel.classList.add("hidden");
  refs.excelPanel.classList.remove("hidden");
  refs.heroPill.textContent = "Offline Single File";
}

async function bootstrapReadonly() {
  if (!IS_READONLY) {
    setStatus("未加载", "badge-idle");
    return;
  }
  try {
    const response = await fetch(window.__SHARE_DATA_URL__ || "./dashboard-data.json");
    const snapshot = await response.json();
    state.dataset = buildDatasetFromNormalizedRows(snapshot.dataset || snapshot);
    if (refs.sourceStatus) {
      refs.sourceStatus.textContent = "分享快照";
      refs.sourceStatus.className = "badge badge-success";
    }
    hydrateDateInput();
    renderDashboard();
  } catch (error) {
    console.error(error);
    showToast("分享版数据加载失败");
  }
}

bootstrapReadonly();
