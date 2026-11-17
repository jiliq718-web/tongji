
// ======================================================
// 1. CẤU HÌNH SUPABASE
// ======================================================

const SUPABASE_URL = "https://zuzyekphhcwwidcbvnqq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_OxHoqPnBjsvYos7WUm2kpg_hk8GEzyd";

const FETCH_SIZE = 500;
const PAGE_SIZE = 200;

const COLS = [
  ["agent_id", ["代理ID", "代理ID", "代理id", "代理编号"]],
  ["report_date", ["日期", "Ngày", "日期日期"]],
  ["daily_salary", ["日薪", "Lương ngày"]],
  ["ad_funds", ["广告资金", "Quỹ quảng cáo", "Quảng cáo"]],
  ["commission", ["佣金", "Hoa hồng"]],
  ["video_promotion", ["视频推广", "Video", "Video quảng bá"]],
  ["total_expense", ["总费用", "Tổng chi phí"]],
  ["level1_deposit", ["一级首存", "一级存", "一级存款", "Cấp 1 gửi"]],
  ["level1_withdrawal", ["一级二存", "一级存二", "一级存款二", "Cấp 1 rút"]],
  ["level2_deposit", ["一级三存", "二级存", "Cấp 2 gửi"]],
  ["total_assets", ["总存款金额", "Tổng tài sản", "Tổng tiền gửi"]],
  ["total_withdrawals", ["总取款金额", "Tổng rút", "Tổng tiền rút"]],
  ["internal_profit", ["内盘盈亏", "Lãi trong nhà"]],
  ["external_profit", ["外盘盈亏", "Lãi ngoài nhà"]],
  ["storage_diff", ["存提差", "Chênh lệch lưu trữ"]],
  ["total_profit", ["总盈亏", "总盈亏额", "Tổng lãi", "Tổng lãi/lỗ"]],
  ["profit_percent", ["利润百分比", "LN %", "Tỷ lệ lợi nhuận"]],
  ["handling_fee", ["手续费", "Phí thủ tục"]],
  ["new_ad_status", ["新广告费是否异常", "Quảng cáo mới", "Tình trạng quảng cáo"]],
  ["cooperation_status", ["合作状态", "Trạng thái hợp tác"]]
];

const NUMERIC_FIELDS = COLS.slice(2, 18).map(c => c[0]);

let db = null;
let currentRows = [];
let hierarchyRows = [];
let statsPage = 1;
let loadingData = false;
let importingExcel = false;

const $ = id => document.getElementById(id);

// ======================================================
// 2. HÀM HỖ TRỢ
// ======================================================

function configured() {
  return SUPABASE_URL.startsWith("https://") &&
    !SUPABASE_ANON_KEY.includes("PASTE_") &&
    SUPABASE_ANON_KEY.length > 20;
}

function msg(id, text, error = false) {
  const el = $(id);
  if (!el) return;
  el.textContent = text;
  el.classList.toggle("error", error);
}

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[：:]/g, "");
}

function cleanId(value) {
  return String(value ?? "").trim();
}

function parseNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }

  const number = Number(
    String(value).replace(/[,，￥¥%\s]/g, "").replace(/[()]/g, "")
  );

  return Number.isFinite(number) ? number : null;
}

function parseDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number" && value > 20000 && value < 80000) {
    const d = XLSX.SSF.parse_date_code(value);

    if (d) {
      return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
    }
  }

  const s = String(value ?? "").trim();

  let m = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);

  if (m) {
    return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  }

  m = s.match(/^(\d{1,2})月(\d{1,2})(?:日|号)?$/);

  if (m) {
    return `${new Date().getFullYear()}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  }

  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})$/);

  if (m) {
    return `${new Date().getFullYear()}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  }

  return s;
}

function dateLabel(value) {
  if (!value) return "";

  const d = new Date(`${value}T00:00:00`);

  return isNaN(d)
    ? value
    : `${d.getMonth() + 1}月${d.getDate()}号`;
}

function fmt(value) {
  if (value === null || value === undefined || value === "") return "—";

  if (typeof value === "number") {
    return value.toLocaleString("vi-VN", {
      maximumFractionDigits: 4
    });
  }

  return String(value);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

// ======================================================
// 3. ĐỌC FILE EXCEL THỐNG KÊ
// ======================================================

function headerIndex(rows) {
  let best = { row: -1, score: 0, map: {} };

  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const cells = (rows[r] || []).map(normalize);
    const map = {};
    let score = 0;

    for (const [key, aliases] of COLS) {
      const idx = cells.findIndex(cell =>
        [key, ...aliases.map(normalize)].includes(cell)
      );

      if (idx >= 0) {
        map[key] = idx;
        score++;
      }
    }

    if (score > best.score) {
      best = { row: r, score, map };
    }
  }

  return best;
}

function rowToRecord(row, map, lastAgent) {
  const out = {};

  for (const [key] of COLS) {
    const idx = map[key];
    out[key] = idx === undefined ? null : row[idx];
  }

  let agent = String(out.agent_id ?? "").trim();

  if (!agent || agent === "undefined" || agent === "null") {
    agent = lastAgent || "";
  }

  if (!agent || ["总计", "tổng"].includes(normalize(agent))) {
    return null;
  }

  out.agent_id = agent;
  out.report_date = parseDate(out.report_date);

  for (const field of NUMERIC_FIELDS) {
    out[field] = parseNumber(out[field]);
  }

  out.new_ad_status = out.new_ad_status == null
    ? ""
    : String(out.new_ad_status).trim();

  out.cooperation_status = out.cooperation_status == null
    ? ""
    : String(out.cooperation_status).trim();

  if (!out.report_date || !/^\d{4}-\d{2}-\d{2}$/.test(out.report_date)) {
    return null;
  }

  return out;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true
  });

  const records = [];
  const skipped = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
      raw: true
    });

    const h = headerIndex(rows);

    if (h.score < 8) {
      skipped.push(sheetName);
      continue;
    }

    let lastAgent = "";

    for (let i = h.row + 1; i < rows.length; i++) {
      const row = rows[i] || [];

      const rawAgent = h.map.agent_id === undefined
        ? ""
        : String(row[h.map.agent_id] ?? "").trim();

      if (
        rawAgent &&
        normalize(rawAgent) !== "总计" &&
        normalize(rawAgent) !== "tổng"
      ) {
        lastAgent = rawAgent;
      }

      const record = rowToRecord(row, h.map, lastAgent);

      if (record) records.push(record);
    }
  }

  if (!records.length) {
    throw new Error("未找到有效数据，请检查Excel中的代理ID和日期标题。");
  }

  const unique = new Map();

  for (const record of records) {
    unique.set(`${record.agent_id}__${record.report_date}`, record);
  }

  if (skipped.length) {
    console.info("已跳过辅助工作表：", skipped);
  }

  return [...unique.values()];
}

// ======================================================
// 4. NHẬP EXCEL THỐNG KÊ - THAY THẾ DỮ LIỆU CŨ
// ======================================================

async function importExcel(file) {
  if (!db) throw new Error("尚未配置Supabase。");
  if (importingExcel) throw new Error("另一个文件正在导入，请稍候。");

  const records = parseWorkbook(await file.arrayBuffer());
  if (!records.length) throw new Error("Excel文件中没有有效数据。");

  const confirmed = confirm(
    `警告：将用新数据替换全部旧数据\n\n` +
    `文件：${file.name}\n` +
    `准备导入：${records.length.toLocaleString("zh-CN")} 条记录\n\n` +
    `旧数据只会在新数据全部上传成功后才被替换。\n` +
    `导入失败时，现有数据保持不变。\n` +
    `代理上下级关系将保留。\n\n确定继续吗？`
  );
  if (!confirmed) {
    msg("dataMessage", "已取消导入，旧数据保持不变。");
    return;
  }

  importingExcel = true;
  const fileInput = $("excelFile");
  if (fileInput) fileInput.disabled = true;

  const BATCH_SIZE = 500;
  const importId = crypto.randomUUID();
  let uploaded = 0;
  let finalized = false;

  try {
    for (let start = 0; start < records.length; start += BATCH_SIZE) {
      const batch = records.slice(start, start + BATCH_SIZE);
      const { error } = await db.rpc("stage_agent_stats", {
        p_import_id: importId,
        p_start_index: start,
        p_rows: batch
      });
      if (error) throw error;

      uploaded = Math.min(start + batch.length, records.length);
      msg(
        "dataMessage",
        `正在上传：${uploaded.toLocaleString("zh-CN")} / ${records.length.toLocaleString("zh-CN")} 条（${Math.floor(uploaded / records.length * 100)}%）`
      );
    }

    msg("dataMessage", "数据上传完成，正在安全替换正式数据，请勿关闭页面...");
    const { data, error } = await db.rpc("finalize_agent_stats", {
      p_import_id: importId,
      p_expected_count: records.length
    });
    if (error) throw error;
    finalized = true;

    currentRows = [];
    hierarchyRows = [];
    await loadData();

    msg(
      "dataMessage",
      `导入成功！已更新 ${Number(data).toLocaleString("zh-CN")} 条记录。文件：${file.name}`
    );
  } catch (error) {
    if (!finalized) {
      try {
        await db.rpc("discard_agent_stats_import", { p_import_id: importId });
      } catch (cleanupError) {
        console.warn("暂存数据清理失败：", cleanupError);
      }
    }
    msg(
      "dataMessage",
      `导入失败：${error.message || error}。正式数据在最终提交前不会被更改；已上传暂存 ${uploaded.toLocaleString("zh-CN")} / ${records.length.toLocaleString("zh-CN")} 条。`,
      true
    );
    throw error;
  } finally {
    importingExcel = false;
    if (fileInput) fileInput.disabled = false;
  }
}

// ======================================================
// 5. TẠO GIAO DIỆN NHẬP LIÊN KẾT HÀNG LOẠT
// ======================================================

function setupBulkHierarchyUI() {
  const view = $("hierarchyView");
  if (!view) return;

  let area = $("bulkHierarchyArea");

  if (!area) {
    area = document.createElement("div");
    area.id = "bulkHierarchyArea";
    area.className = "example-note";
    area.style.marginBottom = "16px";

    area.innerHTML = `
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <button id="downloadHierarchyTemplateBtn" type="button" class="secondary">
          下载代理关系Excel模板
        </button>

        <label class="file-button">
          批量上传代理关系
          <input id="hierarchyExcelFile" type="file" accept=".xlsx,.xls,.csv">
        </label>
      </div>

      <p style="margin-top:10px">
        下载Excel模板，在每一行填写代理ID和直属上级ID，
        然后上传文件即可批量建立代理关系。
      </p>
      <p id="bulkHierarchyMessage" class="message" aria-live="polite"></p>
    `;

    const example = view.querySelector(".example-note");

    if (example) {
      example.insertAdjacentElement("afterend", area);
    } else {
      view.prepend(area);
    }
  }

  $("downloadHierarchyTemplateBtn")?.addEventListener(
    "click",
    downloadHierarchyTemplate
  );

  $("hierarchyExcelFile")?.addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      await importHierarchyExcel(file);
    } catch (error) {
      msg("bulkHierarchyMessage", `导入失败： ${error.message}`, true);
    } finally {
      event.target.value = "";
    }
  });
}

// ======================================================
// 6. TẢI MẪU EXCEL LIÊN KẾT
// ======================================================

function downloadHierarchyTemplate() {
  const rows = [
    ["agent_id", "parent_agent_id"],
    ["6417521", "5290331"],
    ["111853", "118380"],
    ["", ""]
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "上级关系");

  XLSX.writeFile(wb, "代理上级关系模板.xlsx");

  msg("bulkHierarchyMessage", "Excel模板已下载，请将示例ID替换为您自己的ID。");
}

// ======================================================
// 7. NHẬP LIÊN KẾT CẤP TRÊN HÀNG LOẠT
// ======================================================

function parseHierarchyWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true
  });

  const rows = XLSX.utils.sheet_to_json(
    wb.Sheets[wb.SheetNames[0]],
    { header: 1, defval: "" }
  );

  let headerRow = -1;
  let childIndex = -1;
  let parentIndex = -1;

  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const headers = rows[i].map(normalize);

    childIndex = headers.findIndex(v =>
      ["agent_id", "mãđạilý", "mãđạilýcấpdưới", "代理id", "代理编号"].includes(v)
    );

    parentIndex = headers.findIndex(v =>
      ["parent_agent_id", "mãcấp共trựctiếp", "mãcấp共", "上级代理id"].includes(v)
    );

    if (childIndex >= 0 && parentIndex >= 0) {
      headerRow = i;
      break;
    }
  }

  if (headerRow < 0) {
    throw new Error("模板中未找到agent_id和parent_agent_id两列。");
  }

  const result = [];
  const seen = new Set();

  for (let i = headerRow + 1; i < rows.length; i++) {
    const child = cleanId(rows[i][childIndex]);
    const parent = cleanId(rows[i][parentIndex]);

    if (!child && !parent) continue;

    if (!child || !parent) {
      throw new Error(`Dòng ${i + 1}: 缺少代理ID或上级ID。`);
    }

    if (child === parent) {
      throw new Error(`Dòng ${i + 1}: 代理不能将自己设为上级。`);
    }

    if (seen.has(child)) {
      throw new Error(`代理ID ${child} 在文件中重复。`);
    }

    seen.add(child);
    result.push({
      agent_id: child,
      parent_agent_id: parent
    });
  }

  if (!result.length) {
    throw new Error("文件中没有有效的代理关系记录。");
  }

  return result;
}

async function importHierarchyExcel(file) {
  if (!db) throw new Error("Chưa kết nối Supabase.");

  const records = parseHierarchyWorkbook(await file.arrayBuffer());

  const confirmed = confirm(
    `即将建立 ${records.length.toLocaleString("vi-VN")} 个代理。\n\n` +
    `如果代理已有上级，将用新关系更新原关系。\n\n` +
    `是否继续？`
  );

  if (!confirmed) return;

  const button = $("hierarchyExcelFile");
  if (button) button.disabled = true;

  try {
    msg(
      "bulkHierarchyMessage",
      `正在检查并导入 ${records.length.toLocaleString("vi-VN")} 条关系...`
    );

    // Kiểm tra vòng lặp trước khi ghi.
    const map = parentMap();

    for (const item of records) {
      map.set(item.agent_id, item.parent_agent_id);
    }

    for (const item of records) {
      let cursor = item.parent_agent_id;
      const seen = new Set([item.agent_id]);

      while (cursor) {
        if (seen.has(cursor)) {
          throw new Error(`关联 của ${item.agent_id} 会形成循环关系，请检查文件。`);
        }

        seen.add(cursor);
        cursor = map.get(cursor) || "";
      }
    }

    // Ghi theo từng đợt để tránh request quá lớn.
    const batchSize = 200;

    for (let i = 0; i < records.length; i += batchSize) {
      const batch = records.slice(i, i + batchSize);

      const { error } = await db
        .from("agent_hierarchy")
        .upsert(batch, { onConflict: "agent_id" });

      if (error) throw error;

      msg(
        "bulkHierarchyMessage",
        `已保存 ${Math.min(i + batch.length, records.length).toLocaleString("vi-VN")} / ${records.length.toLocaleString("vi-VN")} 条关系...`
      );
    }

    await loadHierarchy();

    msg(
      "bulkHierarchyMessage",
      `完成！已建立关系 ${records.length.toLocaleString("vi-VN")} 个代理。`
    );

  } finally {
    if (button) button.disabled = false;
  }
}

// ======================================================
// 8. QUAN HỆ CẤP TRÊN
// ======================================================

function parentMap() {
  const map = new Map();

  for (const row of hierarchyRows) {
    if (row.agent_id && row.parent_agent_id) {
      map.set(String(row.agent_id), String(row.parent_agent_id));
    }
  }

  return map;
}

function descendantIds(rootId) {
  const root = cleanId(rootId);
  const found = new Set([root]);

  let changed = true;

  while (changed) {
    changed = false;

    for (const row of hierarchyRows) {
      const child = cleanId(row.agent_id);
      const parent = cleanId(row.parent_agent_id);

      if (parent && found.has(parent) && child && !found.has(child)) {
        found.add(child);
        changed = true;
      }
    }
  }

  return found;
}

async function loadHierarchy() {
  if (!db) return;

  try {
    let allRows = [];

    for (let from = 0; ; from += FETCH_SIZE) {
      const { data, error } = await db
        .from("agent_hierarchy")
        .select("agent_id,parent_agent_id")
        .order("agent_id", { ascending: true })
        .range(from, from + FETCH_SIZE - 1);

      if (error) throw error;

      allRows = allRows.concat(data || []);

      if (!data || data.length < FETCH_SIZE) break;
    }

    hierarchyRows = allRows;

    renderHierarchy();
    render();

  } catch (error) {
    msg("hierarchyMessage", `无法加载上级关系： ${error.message}`, true);
  }
}

// ======================================================
// 9. TẠO KHUNG ĐẠI LÝ ĐÃ LIÊN KẾT
// ======================================================

function setupLinkedPanel() {
  const view = $("hierarchyView");
  if (!view || $("linkedBody")) return;

  const panel = document.createElement("div");
  panel.className = "unassigned-panel";
  panel.style.marginTop = "20px";

  panel.innerHTML = `
    <div class="section-heading compact-heading">
      <div>
        <h2>2. 已关联代理列表</h2>
        <p class="muted">以下代理已有直属上级，您可以更新或删除关系。</p>
      </div>
      <strong id="linkedCount" class="count-pill">0 mã</strong>
    </div>

    <div class="table-wrap unassigned-table-wrap">
      <table class="hierarchy-table">
        <thead>
          <tr>
            <th>代理ID</th>
            <th>统计记录数</th>
            <th>直属上级ID</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody id="linkedBody">
          <tr><td colspan="4" class="empty">正在加载...</td></tr>
        </tbody>
      </table>
    </div>
  `;

  const firstPanel = view.querySelector(".unassigned-panel");

  if (firstPanel) {
    firstPanel.insertAdjacentElement("afterend", panel);
  } else {
    view.appendChild(panel);
  }
}

// ======================================================
// 10. HIỂN THỊ HAI DANH SÁCH
// ======================================================

function renderHierarchy() {
  setupLinkedPanel();
  renderAgentLinks();
}

function renderAgentLinks() {
  const unlinkedBody = $("unassignedBody");
  const linkedBody = $("linkedBody");

  if (!unlinkedBody || !linkedBody) return;

  const counts = new Map();

  for (const row of currentRows) {
    const id = cleanId(row.agent_id);
    if (id) counts.set(id, (counts.get(id) || 0) + 1);
  }

  const parents = parentMap();

  const all = [...counts.entries()].sort((a, b) =>
    a[0].localeCompare(b[0], undefined, { numeric: true })
  );

  const query = normalize($("unassignedSearch")?.value || "");

  const filtered = all.filter(([id]) =>
    !query ||
    normalize(id).includes(query) ||
    normalize(parents.get(id) || "").includes(query)
  );

  const unlinked = filtered.filter(([id]) => !parents.has(id));
  const linked = filtered.filter(([id]) => parents.has(id));

  if ($("unassignedCount")) {
    $("unassignedCount").textContent =
      `${unlinked.length.toLocaleString("vi-VN")} mã`;
  }

  if ($("linkedCount")) {
    $("linkedCount").textContent =
      `${linked.length.toLocaleString("vi-VN")} mã`;
  }

  function createAgentRow(id, count, isLinked) {
    const tr = document.createElement("tr");

    const idCell = document.createElement("td");
    idCell.textContent = id;

    const countCell = document.createElement("td");
    countCell.textContent = count.toLocaleString("vi-VN");

    const parentCell = document.createElement("td");
    const input = document.createElement("input");

    input.type = "text";
    input.className = "inline-parent-input";
    input.placeholder = "输入上级代理ID";
    input.value = isLinked ? parents.get(id) : "";

    parentCell.appendChild(input);

    const actionCell = document.createElement("td");
    actionCell.className = "row-actions";

    const saveButton = document.createElement("button");
    saveButton.className = "small-button";
    saveButton.textContent = isLinked ? "更新" : "关联";

    saveButton.addEventListener("click", () => {
      saveParentFor(id, input.value);
    });

    actionCell.appendChild(saveButton);

    if (isLinked) {
      const deleteButton = document.createElement("button");
      deleteButton.className = "danger small-button";
      deleteButton.textContent = "删除关系";

      deleteButton.addEventListener("click", () => {
        deleteParent(id);
      });

      actionCell.appendChild(deleteButton);
    }

    tr.append(idCell, countCell, parentCell, actionCell);
    return tr;
  }

  unlinkedBody.innerHTML = "";

  if (!all.length) {
    unlinkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">暂无统计数据，请先导入Excel。</td></tr>';
  } else if (!unlinked.length) {
    unlinkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">没有尚未关联的代理。</td></tr>';
  } else {
    for (const [id, count] of unlinked) {
      unlinkedBody.appendChild(createAgentRow(id, count, false));
    }
  }

  linkedBody.innerHTML = "";

  if (!all.length) {
    linkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">暂无统计数据。</td></tr>';
  } else if (!linked.length) {
    linkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">暂无已关联的代理。</td></tr>';
  } else {
    for (const [id, count] of linked) {
      linkedBody.appendChild(createAgentRow(id, count, true));
    }
  }
}

// ======================================================
// 11. LƯU LIÊN KẾT CẤP TRÊN
// ======================================================

async function saveParentFor(child, parent) {
  child = cleanId(child);
  parent = cleanId(parent);

  if (!child || !parent) {
    msg("hierarchyMessage", "请填写代理ID和上级ID。", true);
    return;
  }

  if (child === parent) {
    msg("hierarchyMessage", "代理不能将自己设为上级。", true);
    return;
  }

  const map = parentMap();
  let cursor = parent;
  const seen = new Set([child]);

  while (cursor) {
    if (seen.has(cursor)) {
      msg("hierarchyMessage", "关联 này tạo vòng lặp cấp 共. Hãy kiểm tra lại.", true);
      return;
    }

    seen.add(cursor);
    cursor = map.get(cursor) || "";
  }

  // Không gửi updated_at vì bảng hiện tại của bạn chưa có cột này.
  const { error } = await db
    .from("agent_hierarchy")
    .upsert({
      agent_id: child,
      parent_agent_id: parent
    }, {
      onConflict: "agent_id"
    });

  if (error) {
    msg("hierarchyMessage", `无法保存代理关系： ${error.message}`, true);
    return;
  }

  msg("hierarchyMessage", `已关联代理 ${child} 隶属于 ${parent}.`);

  await loadHierarchy();
}

// ======================================================
// 12. XÓA LIÊN KẾT
// ======================================================

async function deleteParent(child) {
  if (!confirm(`确定删除该代理的上级关系： ${child}?`)) return;

  const { error } = await db
    .from("agent_hierarchy")
    .delete()
    .eq("agent_id", child);

  if (error) {
    msg("hierarchyMessage", `无法删除关系： ${error.message}`, true);
    return;
  }

  msg("hierarchyMessage", `已删除该代理的上级关系： ${child}.`);

  await loadHierarchy();
}

// ======================================================
// 13. TẢI TOÀN BỘ DỮ LIỆU THEO TỪNG ĐỢT
// ======================================================

async function loadData() {
  if (!db || loadingData) return;

  loadingData = true;
  msg("dataMessage", "正在加载数据...");

  try {
    let allRows = [];

    for (let from = 0; ; from += FETCH_SIZE) {
      const { data, error } = await db
        .from("agent_stats")
        .select("*")
        .order("id", { ascending: true })
        .range(from, from + FETCH_SIZE - 1);

      if (error) throw error;

      allRows = allRows.concat(data || []);

      msg(
        "dataMessage",
        `正在加载数据： ${allRows.length.toLocaleString("vi-VN")} 条`
      );

      if (!data || data.length < FETCH_SIZE) break;
    }

    currentRows = allRows;
    statsPage = 1;

    render();

    await loadHierarchy();

    msg(
      "dataMessage",
      `Đã tải thành công ${currentRows.length.toLocaleString("vi-VN")} 条数据。`
    );

  } catch (error) {
    msg("dataMessage", `数据加载失败： ${error.message || error}`, true);
    throw error;

  } finally {
    loadingData = false;
  }
}

// ======================================================
// 14. LỌC DỮ LIỆU VÀ CẤP DƯỚI
// ======================================================

function filteredRows() {
  const rawAgent = cleanId($("agentFilter")?.value);
  const agent = normalize(rawAgent);

  const from = $("dateFrom")?.value || "";
  const to = $("dateTo")?.value || "";

  let allowed = null;

  if (rawAgent) {
    const exactExists =
      currentRows.some(r => normalize(r.agent_id) === agent) ||
      hierarchyRows.some(r =>
        normalize(r.agent_id) === agent ||
        normalize(r.parent_agent_id) === agent
      );

    if (exactExists) {
      allowed = descendantIds(rawAgent);
    }
  }

  return currentRows.filter(row => {
    const id = String(row.agent_id ?? "");

    const matchesAgent = !agent ||
      (allowed ? allowed.has(id) : normalize(id).includes(agent));

    return matchesAgent &&
      (!from || row.report_date >= from) &&
      (!to || row.report_date <= to);
  });
}


// Kiểu màu cho 条 tổng từng 个代理 và tổng chung.
(function addTotalRowStyles() {
  if (document.getElementById("agentTotalRowStyles")) return;
  const style = document.createElement("style");
  style.id = "agentTotalRowStyles";
  style.textContent = `
    .agent-total-row td,
    .grand-total-row td {
      background: #ffc000 !important;
      color: #111827 !important;
      font-weight: 700 !important;
      border-top: 2px solid #b8860b !important;
      border-bottom: 2px solid #b8860b !important;
    }
    .agent-total-row td.negative,
    .grand-total-row td.negative {
      color: #d00000 !important;
    }
    .grand-total-row td {
      background: #f5b800 !important;
      border-top: 3px solid #8a6500 !important;
    }
  `;
  document.head.appendChild(style);
})();

function sortedGroupCount(rows) {
  return new Set(rows.map(row => cleanId(row.agent_id)).filter(Boolean)).size;
}

// ======================================================
// 15. HIỂN THỊ THỐNG KÊ
// ======================================================

function render() {
  const rows = filteredRows();

  const body = $("statsBody");
  const foot = $("statsFoot");

  if (!body || !foot) return;

  $("kpiAgents").textContent =
    new Set(rows.map(r => r.agent_id)).size.toLocaleString("vi-VN");

  $("kpiRows").textContent = rows.length.toLocaleString("vi-VN");

  const sum = (items, field) =>
    items.reduce((total, row) => total + (Number(row[field]) || 0), 0);

  $("kpiExpense").textContent = fmt(sum(rows, "total_expense"));
  $("kpiProfit").textContent = fmt(sum(rows, "total_profit"));

  body.innerHTML = "";
  foot.innerHTML = "";

  // Bỏ các bản ghi hoàn toàn trống về mặt số liệu.
  // Video có giá trị 0 vẫn được xem là dữ liệu hợp lệ nếu các cột khác có số.
  const hasActualData = row =>
    NUMERIC_FIELDS.some(field => {
      const value = row[field];
      return value !== null &&
        value !== undefined &&
        value !== "" &&
        Number(value) !== 0;
    }) ||
    Boolean(String(row.new_ad_status || "").trim()) ||
    Boolean(String(row.cooperation_status || "").trim());

  const visibleRows = rows.filter(hasActualData);

  const totalPages = Math.max(1, Math.ceil(sortedGroupCount(visibleRows) / PAGE_SIZE));

  if (statsPage > totalPages) statsPage = totalPages;
  if (statsPage < 1) statsPage = 1;

  if (!visibleRows.length) {
    body.innerHTML =
      '<tr><td colspan="20" class="empty">没有符合筛选条件的数据。</td></tr>';
    updateStatsPagination(0, totalPages);
    return;
  }

  const displayFields = COLS.map(c => c[0]);

  const totalFields = new Set([
    "daily_salary",
    "ad_funds",
    "commission",
    "video_promotion",
    "total_expense",
    "level1_deposit",
    "level1_withdrawal",
    "level2_deposit",
    "total_assets",
    "total_withdrawals",
    "internal_profit",
    "external_profit",
    "storage_diff",
    "total_profit",
    "handling_fee"
  ]);

  // Gom theo mã 个代理, giữ thứ tự mã 个代理 và ngày.
  const groups = new Map();

  for (const row of visibleRows) {
    const id = cleanId(row.agent_id);
    if (!groups.has(id)) groups.set(id, []);
    groups.get(id).push(row);
  }

  // 按每个代理的“总费用”从高到低排序；费用相同时按代理ID排序。
  const sortedGroups = [...groups.entries()].sort((a, b) => {
    const expenseA = sum(a[1], "total_expense");
    const expenseB = sum(b[1], "total_expense");
    return expenseB - expenseA ||
      a[0].localeCompare(b[0], undefined, { numeric: true });
  });

  const pageStart = (statsPage - 1) * PAGE_SIZE;
  const pageGroups = sortedGroups.slice(pageStart, pageStart + PAGE_SIZE);

  function appendDataRow(row) {
    const tr = document.createElement("tr");

    for (const field of displayFields) {
      const td = document.createElement("td");
      let value = row[field];

      if (field === "report_date") {
        value = dateLabel(value);
      } else if (NUMERIC_FIELDS.includes(field)) {
        value = fmt(value);
      } else {
        value = value ?? "";
      }

      td.textContent = value;

      if (
        NUMERIC_FIELDS.includes(field) &&
        typeof row[field] === "number" &&
        row[field] < 0
      ) {
        td.classList.add("negative");
      }

      tr.appendChild(td);
    }

    body.appendChild(tr);
  }

  function appendSubtotal(agentId, agentRows) {
    const tr = document.createElement("tr");
    tr.className = "agent-total-row";

    for (const [i, field] of displayFields.entries()) {
      const td = document.createElement("td");

      if (i === 0) {
        td.textContent = "小计";
      } else if (field === "report_date") {
        td.textContent = `${agentRows.length} 条`;
      } else if (totalFields.has(field)) {
        const number = sum(agentRows, field);
        td.textContent = fmt(number);
        if (number < 0) td.classList.add("negative");
      } else if (field === "profit_percent") {
        const denominator = sum(agentRows, "total_expense");
        td.textContent = denominator
          ? fmt(sum(agentRows, "total_profit") / denominator * 100) + "%"
          : "—";
      } else if (field === "agent_id") {
        td.textContent = agentId;
      } else {
        td.textContent = "";
      }

      tr.appendChild(td);
    }

    body.appendChild(tr);
  }

  for (const [agentId, agentRows] of pageGroups) {
    // Chỉ hiện ngày có dữ liệu; không tạo các 条 ngày rỗng có dấu gạch ngang.
    agentRows.sort((a, b) =>
      String(a.report_date || "").localeCompare(String(b.report_date || ""))
    );

    for (const row of agentRows) {
      appendDataRow(row);
    }

    appendSubtotal(agentId, agentRows);
  }

  // Dòng 小计 CHUNG theo quy tắc:
  // 1) 8 cột hoạt động: cộng tất cả 个代理 trong kết quả lọc.
  // 2) Tổng tài sản, Tổng rút, Phí thủ tục: ưu tiên ID tìm kiếm;
  //    nếu ID đó không có số liệu thì lấy cấp dưới trực tiếp;
  //    nếu cấp 1 cũng không có thì lấy cấp 2.
  // 3) Chênh lệch lưu trữ = Tổng tài sản - Tổng rút.
  // 4) Tổng lãi = Chênh lệch lưu trữ - Tổng chi phí.

  const grandTotal = document.createElement("tr");
  grandTotal.className = "grand-total-row";

  const additiveGrandFields = new Set([
    "daily_salary",
    "ad_funds",
    "commission",
    "video_promotion",
    "total_expense",
    "level1_deposit",
    "level1_withdrawal",
    "level2_deposit"
  ]);

  const searchedAgentId = cleanId($("agentFilter")?.value);
  const hasExactSearchedAgent = searchedAgentId &&
    currentRows.some(row => cleanId(row.agent_id) === searchedAgentId);

  // Lấy dữ liệu theo ngày đang lọc, không bị giới hạn bởi việc tìm cấp dưới.
  const fromDate = $("dateFrom")?.value || "";
  const toDate = $("dateTo")?.value || "";
  const dateRows = currentRows.filter(row =>
    (!fromDate || row.report_date >= fromDate) &&
    (!toDate || row.report_date <= toDate)
  );

  const rootId = hasExactSearchedAgent
    ? searchedAgentId
    : (sortedGroups.length ? sortedGroups[0][0] : "");

  const rootRows = dateRows.filter(row => cleanId(row.agent_id) === rootId);

  const level1Ids = new Set(
    hierarchyRows
      .filter(link => cleanId(link.parent_agent_id) === rootId)
      .map(link => cleanId(link.agent_id))
  );

  const level1Rows = dateRows.filter(row => level1Ids.has(cleanId(row.agent_id)));

  const level2Ids = new Set(
    hierarchyRows
      .filter(link => level1Ids.has(cleanId(link.parent_agent_id)))
      .map(link => cleanId(link.agent_id))
  );

  const level2Rows = dateRows.filter(row => level2Ids.has(cleanId(row.agent_id)));

  function hasFieldData(items, field) {
    return items.some(row => {
      const value = row[field];
      return value !== null &&
        value !== undefined &&
        value !== "" &&
        Number.isFinite(Number(value)) &&
        Number(value) !== 0;
    });
  }

  function fallbackRowsForField(field) {
    if (hasFieldData(rootRows, field)) return rootRows;
    if (hasFieldData(level1Rows, field)) return level1Rows;
    if (hasFieldData(level2Rows, field)) return level2Rows;
    return [];
  }

  let grandAssets = 0;
  let grandWithdrawals = 0;
  let grandHandlingFee = 0;

  for (const [i, field] of displayFields.entries()) {
    const td = document.createElement("td");

    if (i === 0) {
      td.textContent = "小计 CHUNG";
    } else if (field === "report_date") {
      td.textContent = `${visibleRows.length} 条`;
    } else if (additiveGrandFields.has(field)) {
      const number = sum(visibleRows, field);
      td.textContent = fmt(number);
      if (number < 0) td.classList.add("negative");
    } else if (field === "total_assets") {
      grandAssets = sum(fallbackRowsForField("total_assets"), "total_assets");
      td.textContent = fmt(grandAssets);
      if (grandAssets < 0) td.classList.add("negative");
    } else if (field === "total_withdrawals") {
      grandWithdrawals = sum(
        fallbackRowsForField("total_withdrawals"),
        "total_withdrawals"
      );
      td.textContent = fmt(grandWithdrawals);
      if (grandWithdrawals < 0) td.classList.add("negative");
    } else if (field === "handling_fee") {
      grandHandlingFee = sum(
        fallbackRowsForField("handling_fee"),
        "handling_fee"
      );
      td.textContent = fmt(grandHandlingFee);
      if (grandHandlingFee < 0) td.classList.add("negative");
    } else if (field === "storage_diff") {
      const number = grandAssets - grandWithdrawals;
      td.textContent = fmt(number);
      if (number < 0) td.classList.add("negative");
    } else if (field === "total_profit") {
      const grandStorageDiff = grandAssets - grandWithdrawals;
      const grandTotalExpense = sum(visibleRows, "total_expense");
      const number = grandStorageDiff - grandTotalExpense;
      td.textContent = fmt(number);
      if (number < 0) td.classList.add("negative");
    } else if (field === "profit_percent") {
      const grandStorageDiff = grandAssets - grandWithdrawals;
      const grandTotalExpense = sum(visibleRows, "total_expense");
      const grandProfit = grandStorageDiff - grandTotalExpense;
      td.textContent = grandTotalExpense
        ? fmt(grandProfit / grandTotalExpense * 100) + "%"
        : "—";
    } else if (totalFields.has(field)) {
      const number = sum(rootRows, field);
      td.textContent = fmt(number);
      if (number < 0) td.classList.add("negative");
    } else {
      td.textContent = "";
    }

    grandTotal.appendChild(td);
  }

  foot.appendChild(grandTotal);

  updateStatsPagination(sortedGroups.length, totalPages);
}

// ======================================================
// 16. PHÂN TRANG
// ======================================================

function updateStatsPagination(totalRows, totalPages) {
  let box = $("statsPagination");

  if (!box) {
    box = document.createElement("div");
    box.id = "statsPagination";

    box.style.cssText = `
      display:flex;
      align-items:center;
      justify-content:center;
      gap:12px;
      padding:16px;
      flex-wrap:wrap;
    `;

    $("statsTable").parentElement.insertAdjacentElement("afterend", box);
  }

  const start = totalRows ? (statsPage - 1) * PAGE_SIZE + 1 : 0;
  const end = Math.min(statsPage * PAGE_SIZE, totalRows);

  box.innerHTML = `
    <button id="statsPrev" class="secondary" type="button"
      ${statsPage <= 1 ? "disabled" : ""}>上一页</button>

    <span>
      页 ${statsPage} / ${totalPages}
      · 显示代理 ${start}-${end}
      共 ${totalRows.toLocaleString("vi-VN")} 个代理
    </span>

    <button id="statsNext" class="secondary" type="button"
      ${statsPage >= totalPages ? "disabled" : ""}>下一页</button>
  `;

  $("statsPrev").addEventListener("click", () => {
    if (statsPage > 1) {
      statsPage--;
      render();
    }
  });

  $("statsNext").addEventListener("click", () => {
    if (statsPage < totalPages) {
      statsPage++;
      render();
    }
  });
}

// ======================================================
// 17. CHUYỂN TRANG
// ======================================================

function showView(view) {
  $("statsView")?.classList.toggle("hidden", view !== "stats");
  $("hierarchyView")?.classList.toggle("hidden", view !== "hierarchy");

  $("statsTab")?.classList.toggle("active", view === "stats");
  $("hierarchyTab")?.classList.toggle("active", view === "hierarchy");
}

// ======================================================
// 18. ĐĂNG NHẬP VÀ KHỞI ĐỘNG
// ======================================================

function setSignedIn(yes) {
  $("loginPanel")?.classList.toggle("hidden", yes);
  $("appPanel")?.classList.toggle("hidden", !yes);
  $("logoutBtn")?.classList.toggle("hidden", !yes);

  if ($("accountStatus")) {
    $("accountStatus").textContent = yes ? "已登录" : "未登录";
  }
}

async function init() {
  if (!configured()) {
    $("setupNotice")?.classList.remove("hidden");
    msg("loginMessage", "使用前请先在app.js中填写Supabase配置。");
    return;
  }

  db = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  $("setupNotice")?.classList.add("hidden");

  setupBulkHierarchyUI();
  setupLinkedPanel();

  db.auth.onAuthStateChange((_event, session) => {
    setSignedIn(!!session);
  });

  const { data, error } = await db.auth.getSession();

  if (error) {
    msg("loginMessage", error.message, true);
    return;
  }

  const session = data?.session;

  setSignedIn(!!session);

  if (session) {
    try {
      await loadData();
    } catch (error) {
      msg("dataMessage", error.message || String(error), true);
    }
  }
}

// ======================================================
// 19. ĐĂNG NHẬP
// ======================================================

$("loginForm")?.addEventListener("submit", async event => {
  event.preventDefault();

  if (!db) return;

  msg("loginMessage", "正在登录...");

  const { error } = await db.auth.signInWithPassword({
    email: $("email").value.trim(),
    password: $("password").value
  });

  msg(
    "loginMessage",
    error ? error.message : "登录成功。",
    !!error
  );

  if (!error) {
    try {
      await loadData();
    } catch (err) {
      msg("dataMessage", err.message || String(err), true);
    }
  }
});

// ======================================================
// 20. TẠO TÀI KHOẢN
// ======================================================

$("signupBtn")?.addEventListener("click", async () => {
  if (!db) return;

  msg("loginMessage", "正在创建账户...");

  const { data, error } = await db.auth.signUp({
    email: $("email").value.trim(),
    password: $("password").value
  });

  msg(
    "loginMessage",
    error
      ? error.message
      : (
        data.session
          ? "账户创建成功。"
          : "如果项目要求验证，确认邮件已发送。"
      ),
    !!error
  );

  if (!error && data.session) {
    await loadData();
  }
});

// ======================================================
// 21. ĐĂNG XUẤT
// ======================================================

$("logoutBtn")?.addEventListener("click", async () => {
  if (db) await db.auth.signOut();
});

// ======================================================
// 22. NHẬP FILE EXCEL THỐNG KÊ
// ======================================================

$("excelFile")?.addEventListener("change", async event => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    msg("dataMessage", "正在读取Excel文件...");
    await importExcel(file);
  } catch (error) {
    msg("dataMessage", `Excel导入失败： ${error.message || error}`, true);
  } finally {
    event.target.value = "";
  }
});

// ======================================================
// 23. BỘ LỌC
// ======================================================

["agentFilter", "dateFrom", "dateTo"].forEach(id => {
  $(id)?.addEventListener("input", () => {
    statsPage = 1;
    render();
  });
});

$("clearFiltersBtn")?.addEventListener("click", () => {
  $("agentFilter").value = "";
  $("dateFrom").value = "";
  $("dateTo").value = "";

  statsPage = 1;
  render();
});

// ======================================================
// 24. TẢI LẠI VÀ CHUYỂN TRANG
// ======================================================

$("refreshBtn")?.addEventListener("click", async () => {
  try {
    await loadData();
  } catch (error) {
    msg("dataMessage", error.message || String(error), true);
  }
});

$("statsTab")?.addEventListener("click", () => showView("stats"));

$("hierarchyTab")?.addEventListener("click", () => {
  showView("hierarchy");
  renderHierarchy();
});

$("refreshHierarchyBtn")?.addEventListener("click", loadHierarchy);

$("unassignedSearch")?.addEventListener("input", renderAgentLinks);

// ======================================================
// 25. XUẤT EXCEL THỐNG KÊ
// ======================================================

$("exportBtn")?.addEventListener("click", () => {
  const rows = filteredRows();

  if (!rows.length) {
    msg("dataMessage", "没有可导出的数据。", true);
    return;
  }

  const aoa = [COLS.map(c => c[1][0])];

  rows.forEach(row => {
    aoa.push(
      COLS.map(([field]) =>
        field === "report_date"
          ? dateLabel(row[field])
          : row[field] ?? ""
      )
    );
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "ThongKe");

  XLSX.writeFile(wb, "代理统计.xlsx");
});

// ======================================================
// 26. CHẠY WEBSITE
// ======================================================

init();
