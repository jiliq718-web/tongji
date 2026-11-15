
// ======================================================
// 1. CẤU HÌNH SUPABASE
// ======================================================

const SUPABASE_URL = "https://zuzyekphhcwwidcbvnqq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_OxHoqPnBjsvYos7WUm2kpg_hk8GEzyd";

const FETCH_SIZE = 500;
const PAGE_SIZE = 200;

const COLS = [
  ["agent_id", ["代理ID", "Mã đại lý", "代理id", "代理编号"]],
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
    throw new Error("Không tìm thấy dữ liệu hợp lệ. Kiểm tra tiêu đề mã đại lý và ngày trong Excel.");
  }

  const unique = new Map();

  for (const record of records) {
    unique.set(`${record.agent_id}__${record.report_date}`, record);
  }

  if (skipped.length) {
    console.info("Đã bỏ qua sheet phụ:", skipped);
  }

  return [...unique.values()];
}

// ======================================================
// 4. NHẬP EXCEL THỐNG KÊ - THAY THẾ DỮ LIỆU CŨ
// ======================================================

async function importExcel(file) {
  if (!db) throw new Error("Chưa cấu hình Supabase.");
  if (importingExcel) throw new Error("Đang nhập một file khác. Vui lòng chờ.");

  // Đọc và kiểm tra trước khi xóa dữ liệu cũ.
  const records = parseWorkbook(await file.arrayBuffer());

  if (!records.length) {
    throw new Error("File Excel không có dữ liệu hợp lệ.");
  }

  const confirmed = confirm(
    `CẢNH BÁO: THAY THẾ TOÀN BỘ DỮ LIỆU CŨ\n\n` +
    `File mới: ${file.name}\n` +
    `Số dòng chuẩn bị nhập: ${records.length.toLocaleString("vi-VN")}\n\n` +
    `Dữ liệu cũ trong agent_stats sẽ bị thay thế.\n` +
    `Quan hệ cấp trên trong agent_hierarchy vẫn được giữ nguyên.\n\n` +
    `Bạn có chắc chắn muốn tiếp tục không?`
  );

  if (!confirmed) {
    msg("dataMessage", "Đã hủy nhập. Dữ liệu cũ vẫn được giữ nguyên.");
    return;
  }

  importingExcel = true;

  const fileInput = $("excelFile");
  if (fileInput) fileInput.disabled = true;

  try {
    msg("dataMessage", `Đang nhập ${records.length.toLocaleString("vi-VN")} dòng...`);

    // RPC xóa dữ liệu cũ và nhập dữ liệu mới trong cùng giao dịch.
    const { data, error } = await db.rpc("replace_agent_stats", {
      p_rows: records
    });

    if (error) throw error;

    currentRows = [];
    hierarchyRows = [];

    await loadData();

    msg(
      "dataMessage",
      `Thay thế thành công! Đã nhập ${Number(data).toLocaleString("vi-VN")} dòng từ ${file.name}.`
    );

  } catch (error) {
    msg(
      "dataMessage",
      `Nhập thất bại: ${error.message || error}. Nếu RPC lỗi, dữ liệu cũ được giữ nguyên.`,
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
          Tải mẫu Excel liên kết
        </button>

        <label class="file-button">
          Tải lên liên kết hàng loạt
          <input id="hierarchyExcelFile" type="file" accept=".xlsx,.xls,.csv">
        </label>
      </div>

      <p style="margin-top:10px">
        Tải mẫu Excel, điền mã đại lý và mã cấp trên trực tiếp vào từng dòng,
        sau đó tải file lên để liên kết nhiều đại lý cùng lúc.
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
      msg("bulkHierarchyMessage", `Nhập thất bại: ${error.message}`, true);
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

  XLSX.utils.book_append_sheet(wb, ws, "QuanHeCapTren");

  XLSX.writeFile(wb, "mau-lien-ket-cap-tren.xlsx");

  msg("bulkHierarchyMessage", "Đã tải mẫu Excel. Hãy thay các mã ví dụ bằng mã của bạn.");
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
      ["parent_agent_id", "mãcấptrêntrựctiếp", "mãcấptrên", "上级代理id"].includes(v)
    );

    if (childIndex >= 0 && parentIndex >= 0) {
      headerRow = i;
      break;
    }
  }

  if (headerRow < 0) {
    throw new Error("Không tìm thấy hai cột agent_id và parent_agent_id trong file mẫu.");
  }

  const result = [];
  const seen = new Set();

  for (let i = headerRow + 1; i < rows.length; i++) {
    const child = cleanId(rows[i][childIndex]);
    const parent = cleanId(rows[i][parentIndex]);

    if (!child && !parent) continue;

    if (!child || !parent) {
      throw new Error(`Dòng ${i + 1}: Thiếu mã đại lý hoặc mã cấp trên.`);
    }

    if (child === parent) {
      throw new Error(`Dòng ${i + 1}: Đại lý không thể là cấp trên của chính nó.`);
    }

    if (seen.has(child)) {
      throw new Error(`Mã đại lý ${child} bị trùng trong file.`);
    }

    seen.add(child);
    result.push({
      agent_id: child,
      parent_agent_id: parent
    });
  }

  if (!result.length) {
    throw new Error("File không có dòng liên kết hợp lệ.");
  }

  return result;
}

async function importHierarchyExcel(file) {
  if (!db) throw new Error("Chưa kết nối Supabase.");

  const records = parseHierarchyWorkbook(await file.arrayBuffer());

  const confirmed = confirm(
    `Bạn chuẩn bị liên kết ${records.length.toLocaleString("vi-VN")} đại lý.\n\n` +
    `Nếu mã đại lý đã có cấp trên, quan hệ cũ sẽ được cập nhật bằng quan hệ mới.\n\n` +
    `Bạn có muốn tiếp tục không?`
  );

  if (!confirmed) return;

  const button = $("hierarchyExcelFile");
  if (button) button.disabled = true;

  try {
    msg(
      "bulkHierarchyMessage",
      `Đang kiểm tra và nhập ${records.length.toLocaleString("vi-VN")} liên kết...`
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
          throw new Error(`Liên kết của ${item.agent_id} tạo vòng lặp. Hãy kiểm tra file.`);
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
        `Đã lưu ${Math.min(i + batch.length, records.length).toLocaleString("vi-VN")} / ${records.length.toLocaleString("vi-VN")} liên kết...`
      );
    }

    await loadHierarchy();

    msg(
      "bulkHierarchyMessage",
      `Hoàn tất! Đã liên kết ${records.length.toLocaleString("vi-VN")} đại lý.`
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
    msg("hierarchyMessage", `Không tải được quan hệ cấp trên: ${error.message}`, true);
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
        <h2>2. Danh sách đại lý đã liên kết</h2>
        <p class="muted">Các đại lý đã có cấp trên trực tiếp. Có thể cập nhật hoặc xóa liên kết.</p>
      </div>
      <strong id="linkedCount" class="count-pill">0 mã</strong>
    </div>

    <div class="table-wrap unassigned-table-wrap">
      <table class="hierarchy-table">
        <thead>
          <tr>
            <th>Mã đại lý</th>
            <th>Số dòng thống kê</th>
            <th>Mã cấp trên trực tiếp</th>
            <th>Thao tác</th>
          </tr>
        </thead>
        <tbody id="linkedBody">
          <tr><td colspan="4" class="empty">Đang tải...</td></tr>
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
    input.placeholder = "Nhập mã cấp trên";
    input.value = isLinked ? parents.get(id) : "";

    parentCell.appendChild(input);

    const actionCell = document.createElement("td");
    actionCell.className = "row-actions";

    const saveButton = document.createElement("button");
    saveButton.className = "small-button";
    saveButton.textContent = isLinked ? "Cập nhật" : "Liên kết";

    saveButton.addEventListener("click", () => {
      saveParentFor(id, input.value);
    });

    actionCell.appendChild(saveButton);

    if (isLinked) {
      const deleteButton = document.createElement("button");
      deleteButton.className = "danger small-button";
      deleteButton.textContent = "Xóa liên kết";

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
      '<tr><td colspan="4" class="empty">Chưa có dữ liệu thống kê. Hãy nhập Excel trước.</td></tr>';
  } else if (!unlinked.length) {
    unlinkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">Không còn đại lý nào chưa liên kết.</td></tr>';
  } else {
    for (const [id, count] of unlinked) {
      unlinkedBody.appendChild(createAgentRow(id, count, false));
    }
  }

  linkedBody.innerHTML = "";

  if (!all.length) {
    linkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">Chưa có dữ liệu thống kê.</td></tr>';
  } else if (!linked.length) {
    linkedBody.innerHTML =
      '<tr><td colspan="4" class="empty">Chưa có đại lý nào được liên kết.</td></tr>';
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
    msg("hierarchyMessage", "Hãy nhập đầy đủ mã đại lý và mã cấp trên.", true);
    return;
  }

  if (child === parent) {
    msg("hierarchyMessage", "Một đại lý không thể là cấp trên của chính nó.", true);
    return;
  }

  const map = parentMap();
  let cursor = parent;
  const seen = new Set([child]);

  while (cursor) {
    if (seen.has(cursor)) {
      msg("hierarchyMessage", "Liên kết này tạo vòng lặp cấp trên. Hãy kiểm tra lại.", true);
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
    msg("hierarchyMessage", `Không lưu được liên kết: ${error.message}`, true);
    return;
  }

  msg("hierarchyMessage", `Đã liên kết đại lý ${child} trực thuộc ${parent}.`);

  await loadHierarchy();
}

// ======================================================
// 12. XÓA LIÊN KẾT
// ======================================================

async function deleteParent(child) {
  if (!confirm(`Bỏ quan hệ cấp trên của đại lý ${child}?`)) return;

  const { error } = await db
    .from("agent_hierarchy")
    .delete()
    .eq("agent_id", child);

  if (error) {
    msg("hierarchyMessage", `Không xóa được quan hệ: ${error.message}`, true);
    return;
  }

  msg("hierarchyMessage", `Đã bỏ quan hệ cấp trên của ${child}.`);

  await loadHierarchy();
}

// ======================================================
// 13. TẢI TOÀN BỘ DỮ LIỆU THEO TỪNG ĐỢT
// ======================================================

async function loadData() {
  if (!db || loadingData) return;

  loadingData = true;
  msg("dataMessage", "Đang tải dữ liệu...");

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
        `Đang tải dữ liệu: ${allRows.length.toLocaleString("vi-VN")} dòng`
      );

      if (!data || data.length < FETCH_SIZE) break;
    }

    currentRows = allRows;
    statsPage = 1;

    render();

    await loadHierarchy();

    msg(
      "dataMessage",
      `Đã tải thành công ${currentRows.length.toLocaleString("vi-VN")} dòng dữ liệu.`
    );

  } catch (error) {
    msg("dataMessage", `Tải dữ liệu thất bại: ${error.message || error}`, true);
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


// Kiểu màu cho dòng tổng từng đại lý và tổng chung.
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
      '<tr><td colspan="20" class="empty">Không có dữ liệu phù hợp bộ lọc.</td></tr>';
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

  // Gom theo mã đại lý, giữ thứ tự mã đại lý và ngày.
  const groups = new Map();

  for (const row of visibleRows) {
    const id = cleanId(row.agent_id);
    if (!groups.has(id)) groups.set(id, []);
    groups.get(id).push(row);
  }

  const sortedGroups = [...groups.entries()].sort((a, b) =>
    a[0].localeCompare(b[0], undefined, { numeric: true })
  );

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
        td.textContent = "TỔNG";
      } else if (field === "report_date") {
        td.textContent = `${agentRows.length} dòng`;
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
    // Chỉ hiện ngày có dữ liệu; không tạo các dòng ngày rỗng có dấu gạch ngang.
    agentRows.sort((a, b) =>
      String(a.report_date || "").localeCompare(String(b.report_date || ""))
    );

    for (const row of agentRows) {
      appendDataRow(row);
    }

    appendSubtotal(agentId, agentRows);
  }

  // Dòng tổng chung toàn bộ dữ liệu đang lọc, đặt ở cuối bảng.
  const grandTotal = document.createElement("tr");
  grandTotal.className = "grand-total-row";

  for (const [i, field] of displayFields.entries()) {
    const td = document.createElement("td");

    if (i === 0) {
      td.textContent = "TỔNG CHUNG";
    } else if (field === "report_date") {
      td.textContent = `${visibleRows.length} dòng`;
    } else if (totalFields.has(field)) {
      const number = sum(visibleRows, field);
      td.textContent = fmt(number);
      if (number < 0) td.classList.add("negative");
    } else if (field === "profit_percent") {
      const denominator = sum(visibleRows, "total_expense");
      td.textContent = denominator
        ? fmt(sum(visibleRows, "total_profit") / denominator * 100) + "%"
        : "—";
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
      ${statsPage <= 1 ? "disabled" : ""}>Trang trước</button>

    <span>
      Trang ${statsPage} / ${totalPages}
      · Hiển thị đại lý ${start}-${end}
      trên ${totalRows.toLocaleString("vi-VN")} đại lý
    </span>

    <button id="statsNext" class="secondary" type="button"
      ${statsPage >= totalPages ? "disabled" : ""}>Trang sau</button>
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
    $("accountStatus").textContent = yes ? "Đã đăng nhập" : "Chưa đăng nhập";
  }
}

async function init() {
  if (!configured()) {
    $("setupNotice")?.classList.remove("hidden");
    msg("loginMessage", "Hãy điền thông tin Supabase trong app.js trước khi sử dụng.");
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

  msg("loginMessage", "Đang đăng nhập...");

  const { error } = await db.auth.signInWithPassword({
    email: $("email").value.trim(),
    password: $("password").value
  });

  msg(
    "loginMessage",
    error ? error.message : "Đăng nhập thành công.",
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

  msg("loginMessage", "Đang tạo tài khoản...");

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
          ? "Tạo tài khoản thành công."
          : "Đã gửi email xác nhận nếu dự án yêu cầu xác minh."
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
    msg("dataMessage", "Đang đọc file Excel...");
    await importExcel(file);
  } catch (error) {
    msg("dataMessage", `Nhập Excel thất bại: ${error.message || error}`, true);
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
    msg("dataMessage", "Không có dữ liệu để xuất.", true);
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

  XLSX.writeFile(wb, "thong-ke-dai-ly.xlsx");
});

// ======================================================
// 26. CHẠY WEBSITE
// ======================================================

init();
