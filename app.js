const REQUIRED_SHEETS = {
  profit: "领星每日销售数据（手动）",
  plan: "Q2销售预期",
};

const IS_STANDALONE =
  window.__STANDALONE__ === true || window.location.protocol === "file:";
const IS_READONLY = window.__READONLY__ === true;

const state = {
  sourceMode: "excel",
  dataset: null,
  manualDate: "",
  aiProvider: "openai",
  detailFilters: { market: "", country: "", asin: "", product: "", owner: "" },
  weeklyFilters: { market: "", country: "", asin: "", product: "", owner: "" },
  monthlyFilters: { month: "", market: "", country: "", asin: "", owner: "" },
  warningFilters: { risk: "", market: "", country: "", asin: "", owner: "" },
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
  saveGithubTokenBtn: document.getElementById("saveGithubTokenBtn"),
  githubToken: document.getElementById("githubToken"),
  shareStatus: document.getElementById("shareStatus"),
  shareLink: document.getElementById("shareLink"),
  aiStatus: document.getElementById("aiStatus"),
  aiProvider: document.getElementById("aiProvider"),
  aiApiKeyLabel: document.getElementById("aiApiKeyLabel"),
  aiHint: document.getElementById("aiHint"),
  openaiApiKey: document.getElementById("openaiApiKey"),
  saveAiConfigBtn: document.getElementById("saveAiConfigBtn"),
  generateAiBtn: document.getElementById("generateAiBtn"),
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
  warningOwnerFilter: document.getElementById("warningOwnerFilter"),
  detailCount: document.getElementById("detailCount"),
  weeklyCount: document.getElementById("weeklyCount"),
  monthlyCount: document.getElementById("monthlyCount"),
  countryBars: document.getElementById("countryBars"),
  countryTableBody: document.getElementById("countryTableBody"),
  warningTopCards: document.getElementById("warningTopCards"),
  readonlyInfoPanel: document.getElementById("readonlyInfoPanel"),
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

const AI_PROVIDER_META = {
  openai: {
    label: "OpenAI API Key",
    providerName: "OpenAI",
    hint: "当前使用 OpenAI 高质量模式生成当前预警 SKU 的建议。建议先完成飞书同步，再点击生成；发布分享版时会把生成结果一起带上。",
  },
  deepseek: {
    label: "DeepSeek API Key",
    providerName: "DeepSeek",
    hint: "当前使用 DeepSeek 高质量模式生成当前预警 SKU 的建议。建议先完成飞书同步，再点击生成；发布分享版时会把生成结果一起带上。",
  },
  siliconflow: {
    label: "SiliconFlow API Key",
    providerName: "SiliconFlow",
    hint: "当前使用 SiliconFlow 免费模型生成当前预警 SKU 的建议。建议先完成飞书同步，再点击生成；发布分享版时会把生成结果一起带上。",
  },
};

function getAiProviderMeta(provider) {
  return AI_PROVIDER_META[provider] || AI_PROVIDER_META.openai;
}

function applyAiProviderUi(provider, options = {}) {
  const meta = getAiProviderMeta(provider);
  state.aiProvider = provider;
  if (refs.aiProvider) refs.aiProvider.value = provider;
  if (refs.aiApiKeyLabel) refs.aiApiKeyLabel.textContent = meta.label;
  if (refs.aiHint) refs.aiHint.textContent = meta.hint;
  if (refs.openaiApiKey) {
    refs.openaiApiKey.placeholder =
      options.placeholder ||
      `填一次后可直接生成 ${meta.providerName} 高质量 AI 建议`;
    if (options.clearValue) refs.openaiApiKey.value = "";
  }
}

async function fetchJsonOrThrow(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let result = null;
  try {
    result = text ? JSON.parse(text) : null;
  } catch (error) {
    throw new Error(`${url} 返回了非 JSON 内容，请先重启本地仪表盘服务后重试`);
  }
  if (!response.ok) {
    throw new Error(result?.error || `${url} 请求失败`);
  }
  return result;
}

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

if (refs.aiProvider) {
  refs.aiProvider.addEventListener("change", async (event) => {
    const provider = event.target.value || "openai";
    applyAiProviderUi(provider, { clearValue: true });
    await loadAiConfigStatus();
  });
}

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
    resetAiStatus();
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
    const result = await fetchJsonOrThrow(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    state.dataset = buildDatasetFromNormalizedRows(result);
    resetAiStatus();
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

if (refs.saveGithubTokenBtn) {
  refs.saveGithubTokenBtn.addEventListener("click", async () => {
    const token = refs.githubToken?.value?.trim();
    if (!token) {
      showToast("请先填写 GitHub token");
      return;
    }
    try {
      const result = await fetchJsonOrThrow("/api/share/configure", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ githubToken: token }),
      });
      refs.shareStatus.textContent = "已配置";
      refs.shareStatus.className = "badge badge-success";
      if (refs.shareLink && result.publicUrl) refs.shareLink.value = result.publicUrl;
      showToast("GitHub 自动发布配置已保存");
    } catch (error) {
      console.error(error);
      showToast(error.message);
    }
  });
}

if (refs.saveAiConfigBtn) {
  refs.saveAiConfigBtn.addEventListener("click", async () => {
    const apiKey = refs.openaiApiKey?.value?.trim();
    const provider = refs.aiProvider?.value || state.aiProvider || "openai";
    const meta = getAiProviderMeta(provider);
    if (!apiKey) {
      showToast(`请先填写 ${meta.label}`);
      return;
    }
    try {
      const result = await fetchJsonOrThrow("/api/ai/configure", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ provider, apiKey }),
      });
      refs.aiStatus.textContent = "已配置";
      refs.aiStatus.className = "badge badge-success";
      if (refs.openaiApiKey) {
        refs.openaiApiKey.value = "";
        refs.openaiApiKey.placeholder = `已保存，可直接生成 ${meta.providerName} 高质量建议`;
      }
      showToast(`${meta.providerName} 高质量模式已配置（${result.model || "高质量模式"}）`);
    } catch (error) {
      console.error(error);
      showToast(error.message);
    }
  });
}

if (refs.generateAiBtn) {
  refs.generateAiBtn.addEventListener("click", async () => {
    if (!state.dataset) {
      showToast("请先同步飞书或上传 Excel，再生成 AI 建议");
      return;
    }

    try {
      refs.aiStatus.textContent = "生成中";
      refs.aiStatus.className = "badge badge-info";
      const model = buildDashboardModel(state.dataset, state.manualDate);
      if (!model.warningRows.length) {
        refs.aiStatus.textContent = "无预警";
        refs.aiStatus.className = "badge badge-success";
        showToast("当前没有需要生成 AI 建议的预警 SKU");
        return;
      }

      const result = await fetchJsonOrThrow("/api/ai/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          provider: refs.aiProvider?.value || state.aiProvider || "openai",
          effectiveDate: model.effectiveDate,
          compareDate: model.compareDate,
          currentWeekRange: model.currentWeekRange,
          currentMonthLabel: model.currentMonthLabel,
          warningRows: model.warningRows,
        }),
      });

      state.dataset = {
        ...state.dataset,
        aiInsights: result.aiInsights,
      };
      refs.aiStatus.textContent = "已生成";
      refs.aiStatus.className = "badge badge-success";
      renderDashboard();
      showToast(`已生成 ${result.aiInsights?.suggestions?.length || 0} 条高质量 AI 建议`);
    } catch (error) {
      console.error(error);
      refs.aiStatus.textContent = "生成失败";
      refs.aiStatus.className = "badge badge-danger";
      showToast(error.message);
    }
  });
}

if (refs.publishShareBtn) {
  refs.publishShareBtn.addEventListener("click", async () => {
    if (!state.dataset) {
      showToast("请先同步飞书或上传 Excel，再发布分享版");
      return;
    }
    try {
      refs.shareStatus.textContent = "发布中";
      refs.shareStatus.className = "badge badge-info";
      const result = await fetchJsonOrThrow("/api/share/publish", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dataset: state.dataset,
          publishedAt: new Date().toISOString(),
        }),
      });
      refs.shareStatus.textContent = "已发布";
      refs.shareStatus.className = "badge badge-success";
      refs.shareLink.value = result.publicUrl || result.localUrl || "";
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
  const dailySummarySheet = workbook.Sheets["日毛利汇总看板"];
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
  const dailySummaryRows = dailySummarySheet
    ? normalizeDailySummaryRows(
        XLSX.utils.sheet_to_json(dailySummarySheet, {
          defval: null,
          raw: false,
        })
      )
    : [];

  return {
    profitRows,
    planRows,
    dailySummaryRows,
    ownerRows,
    weeklyTrackRows,
    monthlyTrackRows,
    aiInsights: workbook.AI_INSIGHTS || null,
    sourceMeta: { source: "excel" },
  };
}

function buildDatasetFromNormalizedRows(payload) {
  return {
    profitRows: normalizeProfitRows(payload.profitRows || []),
    planRows: normalizePlanRows(payload.planRows || []),
    dailySummaryRows: normalizeDailySummaryRows(payload.dailySummaryRows || []),
    ownerRows: normalizeOwnerRows(payload.ownerRows || []),
    weeklyTrackRows: normalizeWeeklyTrackRows(payload.weeklyTrackRows || []),
    monthlyTrackRows: normalizeMonthlyTrackRows(payload.monthlyTrackRows || []),
    aiInsights: payload.aiInsights || null,
    sourceMeta: {
      source: payload.sourceMeta?.source || "feishu",
      updatedAt:
        payload.updatedAt ||
        payload.sourceMeta?.updatedAt ||
        new Date().toISOString(),
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
      month: normalizeText(row["月份"] ?? row.month),
      weekNo: normalizeInteger(row["周序号"] ?? row.weekNo),
      adSpend: normalizeNumber(row["广告费"] ?? row["广告花费"] ?? row.adSpend),
      grossProfitRmb: normalizeNumber(
        row["订单毛利润"] ??
          row["今日毛利RMB"] ??
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

function normalizeDailySummaryRows(rows) {
  return rows
    .map((row) => ({
      date: normalizeDate(row["日期"] ?? row.date),
          previousDate: normalizeDate(row["昨日日期"] ?? row.previousDate),
      country: normalizeText(row["国家"] ?? row.country),
      asin: normalizeText(row["ASIN"] ?? row["asin"] ?? row.asin),
      product: normalizeText(row["品名"] ?? row.product),
      marketLevel: normalizeText(row["市场等级"] ?? row.marketLevel),
      listingLevel: normalizeText(row["链接分级"] ?? row.listingLevel),
      owner: normalizeText(row["负责人"] ?? row.owner),
      latest: normalizeNumber(row["今日毛利RMB"] ?? row.latest),
      previous: normalizeNumber(row["昨日毛利RMB"] ?? row.previous),
      diff: normalizeNumber(row["环比变化"] ?? row.diff),
      diffPct:
        row["环比百分比"] === null || row["环比百分比"] === undefined
          ? null
          : normalizeNumber(row["环比百分比"]),
          dailyPlan: normalizeNumber(row["单日利润预期"] ?? row["单日预期RMB"] ?? row.dailyPlan),
          vsPlan: normalizeNumber(row["vs预期"] ?? row.vsPlan),
          adSpend: normalizeNumber(row["今日广告花费"] ?? row["广告花费"] ?? row.adSpend),
          adSpendPlan: normalizeNumber(row["广告花费预期"] ?? row.adSpendPlan),
          adVsPlan: normalizeNumber(row["广告vs预期"] ?? row.adVsPlan),
      status: normalizeText(row["状态"] ?? row.status),
      weekNo: normalizeInteger(row["周序号"] ?? row.weekNo),
    }))
    .filter((row) => row.date && row.country && row.asin);
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
      adSpend: normalizeNumber(row["广告花费"] ?? row.adSpend),
      adSpendPlan: normalizeNumber(row["广告花费预期"] ?? row.adSpendPlan),
      adSpendProgress:
        row["广告花费进度"] === null || row["广告花费进度"] === undefined
          ? null
          : normalizeNumber(row["广告花费进度"]),
      completion: normalizeCompletion(
        row["完成率"],
        row["周毛利预期RMB"] ?? row.weeklyPlan,
        row["周毛利实际RMB"] ?? row.weeklyActual
      ),
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
      adSpend: normalizeNumber(row["广告花费"] ?? row.adSpend),
      adSpendPlan: normalizeNumber(row["广告花费预期"] ?? row.adSpendPlan),
      adSpendProgress:
        row["广告花费进度"] === null || row["广告花费进度"] === undefined
          ? null
          : normalizeNumber(row["广告花费进度"]),
      completion: normalizeCompletion(
        row["完成率"],
        row["月毛利预期RMB"] ?? row.monthlyPlan,
        row["月毛利实际RMB"] ?? row.monthlyActual
      ),
      gap: normalizeNumber(row["差异"] ?? row.gap),
    }))
    .filter((row) => row.country && row.asin);
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "object") {
    if (value.text !== undefined) value = value.text;
    else if (value.value !== undefined) value = value.value;
    else if (value.name !== undefined) value = value.name;
  }
  const cleaned = String(value).replace(/,/g, "").trim();
  if (cleaned.endsWith("%")) {
    const parsedPct = Number(cleaned.replace("%", ""));
    return Number.isFinite(parsedPct) ? parsedPct / 100 : 0;
  }
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeCompletion(value, planValue, actualValue) {
  if (value !== null && value !== undefined && value !== "") {
    const direct = normalizeNumber(value);
    if (Number.isFinite(direct)) return direct;
  }
  const plan = normalizeNumber(planValue);
  const actual = normalizeNumber(actualValue);
  if (!plan) return null;
  return actual / plan;
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
  renderReadonlyInfo(model);
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
  const yesterday = getLocalDateOffset(-1);
  const suggestedDate =
    [...availableDates].reverse().find((date) => date <= yesterday) || latestDate;
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

  const dailySummarySource = (dataset.dailySummaryRows || []).filter(
    (row) => row.date === effectiveDate
  );
  const detailRows = (
    dailySummarySource.length
      ? dailySummarySource.map((row) => ({
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
          latest: round(row.latest),
          previous: round(row.previous),
          diff: round(row.diff),
          diffPct: row.diffPct,
          dailyPlan: round(row.dailyPlan),
          vsPlan: round(row.vsPlan),
          adSpend: round(row.adSpend),
          adSpendPlan: round(row.adSpendPlan),
          adVsPlan: round(row.adVsPlan),
          status: row.status || "",
        }))
      : [...groupedPlanRows.values()]
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
              adSpend: round(
                shoeDryerProfitRows
                  .filter(
                    (item) =>
                      item.date === effectiveDate &&
                      item.country === row.country &&
                      item.asin === row.asin
                  )
                  .reduce((sum, item) => sum + (item.adSpend || 0), 0)
              ),
              adSpendPlan: 0,
              adVsPlan: 0,
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
    )
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
          adSpend: round(row.adSpend || 0),
          adSpendPlan: round(row.adSpendPlan || 0),
          adSpendProgress: row.adSpendProgress,
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
            adSpend: 0,
            adSpendPlan: 0,
            adSpendProgress: null,
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
    adSpend: round(row.adSpend || 0),
    adSpendPlan: round(row.adSpendPlan || 0),
    adSpendProgress: row.adSpendProgress,
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
  const aiSuggestionMap =
    dataset.aiInsights?.effectiveDate === effectiveDate
      ? new Map(
          (dataset.aiInsights?.suggestions || []).map((item) => [item.key, item])
        )
      : new Map();

  const warningRows = buildWarningRows(
    detailRows,
    weeklyRows,
    shoeDryerProfitRows,
    effectiveDate,
    aiSuggestionMap
  );
  const filteredWarningRows = warningRows.filter((row) => {
    if (state.warningFilters.risk && row.riskLevelLabel !== state.warningFilters.risk) return false;
    if (state.warningFilters.market && row.marketLevel !== state.warningFilters.market) return false;
    if (state.warningFilters.country && row.country !== state.warningFilters.country) return false;
    if (state.warningFilters.asin && row.asin !== state.warningFilters.asin) return false;
    if (state.warningFilters.owner && (row.owner || "") !== state.warningFilters.owner) return false;
    return true;
  });
  const sortedWarningRows = sortRows(filteredWarningRows, state.warningSort);

  return {
    sourceMeta: dataset.sourceMeta,
    aiInsights: dataset.aiInsights || null,
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
      '<tr><td colspan="16" class="table-placeholder">当前没有需要重点处理的预警 SKU</td></tr>';
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
        <td class="${toneClass(row.adSpend)}">${formatCurrency(row.adSpend)}</td>
        <td class="${toneClass(row.adVsPlan)}">${formatCurrency(row.adVsPlan)}</td>
        <td class="${toneClass(row.weeklyCompletion)}">${formatPercent(row.weeklyCompletion)}</td>
        <td>${row.issueTypes.join(" / ")}</td>
        <td>${row.aiAdvice}</td>
        <td>${row.priority}</td>
      </tr>
    `
    )
    .join("");
}

function renderReadonlyInfo(model) {
  if (!refs.readonlyInfoPanel) return;
  refs.readonlyInfoPanel.classList.remove("empty-state");
  const updatedAt = model.sourceMeta?.updatedAt
    ? new Date(model.sourceMeta.updatedAt).toLocaleString("zh-CN", { hour12: false })
    : "未知";
  const aiGeneratedAt = model.aiInsights?.generatedAt
    ? new Date(model.aiInsights.generatedAt).toLocaleString("zh-CN", { hour12: false })
    : "";
  refs.readonlyInfoPanel.innerHTML = `
    <article class="highlight-card blue">
      <h4>数据更新时间</h4>
      <strong>${updatedAt}</strong>
      <span>当前页面展示的是最近一次发布的只读快照</span>
    </article>
    <article class="highlight-card accent">
      <h4>页面说明</h4>
      <span>该链接只支持查看、筛选和排序，不支持飞书同步与上传 Excel。${aiGeneratedAt ? ` AI建议生成时间：${aiGeneratedAt}` : ""}</span>
    </article>
    <article class="highlight-card good">
      <h4>筛选提示</h4>
      <span>可按风险等级、市场等级、月份、国家、ASIN、负责人等筛选，并点击表头对关键数值排序。</span>
    </article>
  `;
}

function renderDetailTable(rows, summary) {
  refs.detailCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.detailTableBody.innerHTML =
      '<tr><td colspan="15" class="table-placeholder">没有 ASIN 明细</td></tr>';
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
        <td class="${toneClass(row.adSpend)}">${formatCurrency(row.adSpend)}</td>
        <td class="${toneClass(row.adVsPlan)}">${formatCurrency(row.adVsPlan)}</td>
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
        <td class="${toneClass(row.adSpend)}">${formatCurrency(row.adSpend)}</td>
        <td class="${toneClass(row.adSpendPlan)}">${formatCurrency(row.adSpendPlan)}</td>
        <td class="${toneClass(row.adSpendProgress)}">${formatPercent(row.adSpendProgress)}</td>
      </tr>
    `
    )
    .join("") + renderWeeklySummaryRow(summary);
}

function renderMonthlyTable(rows, summary) {
  refs.monthlyCount.textContent = `${rows.length} 条`;
  if (!rows.length) {
    refs.monthlyTableBody.innerHTML =
      '<tr><td colspan="14" class="table-placeholder">没有月进度数据</td></tr>';
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
        <td class="${toneClass(row.adSpend)}">${formatCurrency(row.adSpend)}</td>
        <td class="${toneClass(row.adSpendPlan)}">${formatCurrency(row.adSpendPlan)}</td>
        <td class="${toneClass(row.adSpendProgress)}">${formatPercent(row.adSpendProgress)}</td>
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
  const adSpend = round(rows.reduce((sum, row) => sum + (row.adSpend || 0), 0));
  const adVsPlan = round(rows.reduce((sum, row) => sum + (row.adVsPlan || 0), 0));
  return {
    latest,
    previous,
    diff,
    diffPct: calculateTrendRate(diff, previous),
    dailyPlan,
    vsPlan,
    adSpend,
    adVsPlan,
  };
}

function buildWeeklySummary(rows) {
  const weeklyPlan = round(rows.reduce((sum, row) => sum + (row.weeklyPlan || 0), 0));
  const weeklyActual = round(rows.reduce((sum, row) => sum + (row.weeklyActual || 0), 0));
  const adSpend = round(rows.reduce((sum, row) => sum + (row.adSpend || 0), 0));
  const adSpendPlan = round(rows.reduce((sum, row) => sum + (row.adSpendPlan || 0), 0));
  const gap = round(rows.reduce((sum, row) => sum + (row.gap || 0), 0));
  return {
    weeklyPlan,
    weeklyActual,
    adSpend,
    adSpendPlan,
    adSpendProgress: adSpendPlan === 0 ? null : adSpend / adSpendPlan,
    completion: weeklyPlan === 0 ? 0 : weeklyActual / weeklyPlan,
    gap,
  };
}

function buildWarningKey(row) {
  return `${row.country || ""}__${row.asin || ""}__${row.product || ""}`;
}

function buildWarningRows(detailRows, weeklyRows, profitRows, effectiveDate, aiSuggestionMap = new Map()) {
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

      const warningKey = buildWarningKey(row);
      const riskLevelLabel =
        riskScore >= 6 ? "高风险" : riskScore >= 3 ? "中风险" : "低风险";
      const riskLevelClass =
        riskScore >= 6 ? "warn-high" : riskScore >= 3 ? "warn-mid" : "warn-low";
      const priority = riskScore >= 6 ? "今日处理" : riskScore >= 3 ? "本周优先" : "持续观察";
      const generated = aiSuggestionMap.get(warningKey);
      const aiSummary = generated?.summary || buildAiSummary(row, weekly, issueTypes);
      const aiAdvice = generated?.advice || buildAiAdvice(row, weekly, issueTypes);

      return {
        ...row,
        warningKey,
        marketLevel: row.marketLevel || "",
        listingLevel: row.listingLevel || "",
        weeklyPlan: weekly.weeklyPlan || 0,
        weeklyActual: weekly.weeklyActual || 0,
        weeklyCompletion: weekly.completion || 0,
        issueTypes,
        riskScore,
        riskLevelLabel,
        riskLevelClass,
        aiSummary,
        aiAdvice,
        priority,
        aiProvider: generated ? "deepseek" : "rules",
        followOwner: row.owner || "待分配",
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.riskScore - a.riskScore || a.country.localeCompare(b.country));
}

function buildAiSummary(row, weekly, issueTypes) {
  if (issueTypes.includes("负毛利") && issueTypes.includes("低于预期")) {
    return "今日毛利转负且未达预期，属于需要优先处理的经营异常。";
  }
  if (issueTypes.includes("高预期低达成")) {
    return "当前目标高但达成偏低，说明放量节奏与转化承接不匹配。";
  }
  if (issueTypes.includes("连续下滑")) {
    return "最近 3 天毛利连续走弱，说明问题不是单日偶发波动。";
  }
  if (issueTypes.includes("异常波动")) {
    return "当日波动幅度异常，建议优先排查价格、广告或活动扰动。";
  }
  return `当前主要问题为${issueTypes.join("、")}，建议负责人尽快跟进。`;
}

function buildAiAdvice(row, weekly, issueTypes) {
  const actions = [];
  const countryHint = row.country ? `${row.country}站` : "当前站点";
  const listingHint = row.listingLevel ? `${row.listingLevel}级链接` : "当前链接";
  const marketHint = row.marketLevel ? `${row.marketLevel}级市场` : "当前市场";

  if (issueTypes.includes("负毛利")) {
    actions.push(`先检查${countryHint}${row.asin}的广告花费、折扣和定价，确认是否存在亏损投放或异常优惠`);
  }
  if (issueTypes.includes("低于预期")) {
    actions.push(`对比单日预期与实际毛利差额，优先看${listingHint}的点击率、转化率和是否需要补量`);
  }
  if (issueTypes.includes("连续下滑")) {
    actions.push("连续 2-3 天跟踪 Sessions、CVR、广告占比和排名变化，判断是否为持续转弱而非单日扰动");
  }
  if (issueTypes.includes("高预期低达成")) {
    actions.push(`本周完成率仅 ${formatPercent(weekly.completion || 0)}，说明${marketHint}的放量节奏偏慢，建议优先调整投放和活动安排`);
  }
  if (issueTypes.includes("异常波动")) {
    actions.push("检查是否存在价格切换、Coupon/促销、异常广告放量或库存波动导致的单日失真");
  }
  if (!actions.length) {
    actions.push("建议继续观察未来 1-2 天毛利、广告花费与转化变化，再决定是否调整");
  }
  return actions.join("；");
}

function buildMonthlySummary(rows) {
  const monthlyPlan = round(rows.reduce((sum, row) => sum + (row.monthlyPlan || 0), 0));
  const monthlyActual = round(rows.reduce((sum, row) => sum + (row.monthlyActual || 0), 0));
  const adSpend = round(rows.reduce((sum, row) => sum + (row.adSpend || 0), 0));
  const adSpendPlan = round(rows.reduce((sum, row) => sum + (row.adSpendPlan || 0), 0));
  const gap = round(rows.reduce((sum, row) => sum + (row.gap || 0), 0));
  return {
    monthlyPlan,
    monthlyActual,
    adSpend,
    adSpendPlan,
    adSpendProgress: adSpendPlan === 0 ? null : adSpend / adSpendPlan,
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
      <td class="${toneClass(summary.adSpend)}">${formatCurrency(summary.adSpend)}</td>
      <td class="${toneClass(summary.adVsPlan)}">${formatCurrency(summary.adVsPlan)}</td>
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
      <td class="${toneClass(summary.adSpend)}">${formatCurrency(summary.adSpend)}</td>
      <td class="${toneClass(summary.adSpendPlan)}">${formatCurrency(summary.adSpendPlan)}</td>
      <td class="${toneClass(summary.adSpendProgress)}">${formatPercent(summary.adSpendProgress)}</td>
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
      <td class="${toneClass(summary.adSpend)}">${formatCurrency(summary.adSpend)}</td>
      <td class="${toneClass(summary.adSpendPlan)}">${formatCurrency(summary.adSpendPlan)}</td>
      <td class="${toneClass(summary.adSpendProgress)}">${formatPercent(summary.adSpendProgress)}</td>
    </tr>
  `;
}

function hydrateFilters(model) {
  populateFilter(refs.warningRiskFilter, uniqueValues(model.warningRows, "riskLevelLabel"), "全部风险等级", state.warningFilters.risk);
  populateFilter(refs.warningMarketFilter, uniqueValues(model.warningRows, "marketLevel"), "全部市场等级", state.warningFilters.market);
  populateFilter(refs.warningCountryFilter, uniqueValues(model.warningRows, "country"), "全部国家", state.warningFilters.country);
  populateFilter(refs.warningAsinFilter, uniqueValues(model.warningRows, "asin"), "全部 ASIN", state.warningFilters.asin);
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

  const spotChecks = [
    ["美国", "2026-04-22", 2003.42],
    ["美国", "2026-04-21", -453.49],
  ];
  if (model.effectiveDate === "2026-04-22" && model.compareDate === "2026-04-21") {
    const byCountry = new Map(model.countrySummary.map((row) => [row.country, row]));
    for (const [country, date, expected] of spotChecks) {
      const target = byCountry.get(country);
      const actual = date === model.effectiveDate ? target?.latest : target?.previous;
      if (round(actual || 0) !== round(expected)) {
        issues.push(`对账失败：${country} ${date} 预期 ${expected}，实际 ${actual}`);
      }
    }
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
  if (!state.manualDate && availableDates.length) {
    const yesterday = getLocalDateOffset(-1);
    refs.manualDate.value =
      [...availableDates].reverse().find((date) => date <= yesterday) ||
      availableDates[availableDates.length - 1];
  }
}

function getLocalDateOffset(offsetDays) {
  const now = new Date();
  const local = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() + offsetDays,
    12,
    0,
    0
  );
  const y = local.getFullYear();
  const m = String(local.getMonth() + 1).padStart(2, "0");
  const d = String(local.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function setStatus(label, className) {
  refs.sourceStatus.textContent = label;
  refs.sourceStatus.className = `badge ${className}`;
}

function resetAiStatus() {
  if (!refs.aiStatus) return;
  refs.aiStatus.textContent = "未生成";
  refs.aiStatus.className = "badge badge-idle";
}

let toastTimer = null;
function showToast(message) {
  refs.toast.textContent = message;
  refs.toast.classList.remove("hidden");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => refs.toast.classList.add("hidden"), 3200);
}

async function loadAiConfigStatus() {
  if (IS_READONLY || IS_STANDALONE || !refs.aiStatus) return;
  try {
    const provider = refs.aiProvider?.value || state.aiProvider || "openai";
    const meta = getAiProviderMeta(provider);
    const result = await fetchJsonOrThrow(`/api/ai/configure?provider=${encodeURIComponent(provider)}`);
    applyAiProviderUi(result.provider || provider);
    if (result.configured) {
      refs.aiStatus.textContent = "已配置";
      refs.aiStatus.className = "badge badge-success";
      if (refs.openaiApiKey) {
        refs.openaiApiKey.placeholder = `已保存，可直接生成（${result.model || meta.providerName + " 高质量模式"}）`;
      }
    } else {
      refs.aiStatus.textContent = "未生成";
      refs.aiStatus.className = "badge badge-idle";
      if (refs.openaiApiKey) {
        refs.openaiApiKey.placeholder = `填一次后可直接生成 ${meta.providerName} 高质量 AI 建议`;
      }
    }
  } catch (error) {
    console.warn("loadAiConfigStatus failed", error);
  }
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

applyAiProviderUi(state.aiProvider);

async function bootstrapReadonly() {
  if (!IS_READONLY) {
    setStatus("未加载", "badge-idle");
    return;
  }
  try {
    document.body.classList.add("readonly-viewer");
    const snapshot = await fetchJsonOrThrow(window.__SHARE_DATA_URL__ || "./dashboard-data.json");
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

loadAiConfigStatus();
bootstrapReadonly();
