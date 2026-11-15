
// ======================================================
// 1. CẤU HÌNH SUPABASE
// ======================================================

const SUPABASE_URL = "https://zuzyekphhcwwidcbvnqq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_OxHoqPnBjsvYos7WUm2kpg_hk8GEzyd";

// ======================================================
// 2. CẤU HÌNH DỮ LIỆU
// ======================================================

const COLS = [
  ["agent_id", ["代理ID","Mã đại lý","代理id","代理编号"]],
  ["report_date", ["日期","Ngày","日期日期"]],
  ["daily_salary", ["日薪","Lương ngày"]],
  ["ad_funds", ["广告资金","Quỹ quảng cáo","Quảng cáo"]],
  ["commission", ["佣金","Hoa hồng"]],
  ["video_promotion", ["视频推广","Video","Video quảng bá"]],
  ["total_expense", ["总费用","Tổng chi phí"]],
  ["level1_deposit", ["一级首存","一级存","一级存款","Cấp 1 gửi"]],
  ["level1_withdrawal", ["一级二存","一级存二","一级存款二","Cấp 1 rút"]],
  ["level2_deposit", ["一级三存","二级存","Cấp 2 gửi"]],
  ["total_assets", ["总存款金额","Tổng tài sản","Tổng tiền gửi"]],
  ["total_withdrawals", ["总取款金额","Tổng rút","Tổng tiền rút"]],
  ["internal_profit", ["内盘盈亏","Lãi trong nhà"]],
  ["external_profit", ["外盘盈亏","Lãi ngoài nhà"]],
  ["storage_diff", ["存提差","Chênh lệch lưu trữ"]],
  ["total_profit", ["总盈亏","总盈亏额","Tổng lãi","Tổng lãi/lỗ"]],
  ["profit_percent", ["利润百分比","LN %","Tỷ lệ lợi nhuận"]],
  ["handling_fee", ["手续费","Phí thủ tục"]],
  ["new_ad_status", ["新广告费是否异常","Quảng cáo mới","Tình trạng quảng cáo"]],
  ["cooperation_status", ["合作状态","Trạng thái hợp tác"]]
];

const NUMERIC_FIELDS = COLS.slice(2,18).map(x => x[0]);

// Chỉ lấy 500 dòng mỗi lần từ Supabase
const FETCH_SIZE = 500;

// Chỉ hiển thị 200 dòng trên mỗi trang
const PAGE_SIZE = 200;

let db = null;
let currentRows = [];
let hierarchyRows = [];
let statsPage = 1;
let loadingData = false;

const $ = id => document.getElementById(id);

// ======================================================
// 3. HÀM HỖ TRỢ
// ======================================================

function configured(){
  return SUPABASE_URL.startsWith("https://") &&
    !SUPABASE_URL.includes("PASTE_") &&
    !SUPABASE_ANON_KEY.includes("PASTE_") &&
    SUPABASE_ANON_KEY.length > 20;
}

function msg(id, text, error = false){
  const el = $(id);
  if(!el) return;

  el.textContent = text;
  el.classList.toggle("error", error);
}

function normalize(s){
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[：:]/g, "");
}

function cleanId(v){
  return String(v ?? "").trim();
}

function parseNumber(v){
  if(typeof v === "number"){
    return Number.isFinite(v) ? v : null;
  }

  if(v === null || v === undefined || String(v).trim() === ""){
    return null;
  }

  const s = String(v)
    .replace(/[,，￥¥%\s]/g, "")
    .replace(/[()]/g, "");

  const n = Number(s);

  return Number.isFinite(n) ? n : null;
}

function parseDate(v){
  if(v instanceof Date && !isNaN(v)){
    return v.toISOString().slice(0,10);
  }

  if(typeof v === "number" && v > 20000 && v < 80000){
    const d = XLSX.SSF.parse_date_code(v);

    if(d){
      return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    }
  }

  const s = String(v ?? "").trim();

  let m = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);

  if(m){
    return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  }

  m = s.match(/^(\d{1,2})月(\d{1,2})(?:日|号)?$/);

  if(m){
    return `${new Date().getFullYear()}-${m[1].padStart(2,"0")}-${m[2].padStart(2,"0")}`;
  }

  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})$/);

  if(m){
    return `${new Date().getFullYear()}-${m[1].padStart(2,"0")}-${m[2].padStart(2,"0")}`;
  }

  return s;
}

function dateLabel(s){
  if(!s) return "";

  const d = new Date(`${s}T00:00:00`);

  return isNaN(d)
    ? s
    : `${d.getMonth()+1}月${d.getDate()}号`;
}

function fmt(v){
  if(v === null || v === undefined || v === ""){
    return "—";
  }

  if(typeof v === "number"){
    return v.toLocaleString("vi-VN", {
      maximumFractionDigits: 4
    });
  }

  return String(v);
}

// ======================================================
// 4. ĐỌC FILE EXCEL
// ======================================================

function headerIndex(rows){
  let best = {row:-1, score:0, map:{}};

  for(let r = 0; r < Math.min(rows.length,20); r++){
    const cells = (rows[r] || []).map(normalize);
    const map = {};
    let score = 0;

    for(const [key,aliases] of COLS){
      const idx = cells.findIndex(c =>
        [key, ...aliases.map(normalize)].includes(c)
      );

      if(idx >= 0){
        map[key] = idx;
        score++;
      }
    }

    if(score > best.score){
      best = {row:r, score, map};
    }
  }

  return best;
}

function rowToRecord(row,map,lastAgent){
  const out = {};

  for(const [key] of COLS){
    const idx = map[key];
    out[key] = idx === undefined ? null : row[idx];
  }

  let agent = String(out.agent_id ?? "").trim();

  if(!agent || agent === "undefined" || agent === "null"){
    agent = lastAgent || "";
  }

  if(!agent || normalize(agent) === "总计" || normalize(agent) === "tổng"){
    return null;
  }

  out.agent_id = agent;
  out.report_date = parseDate(out.report_date);

  for(const f of NUMERIC_FIELDS){
    out[f] = parseNumber(out[f]);
  }

  out.new_ad_status = out.new_ad_status == null
    ? ""
    : String(out.new_ad_status).trim();

  out.cooperation_status = out.cooperation_status == null
    ? ""
    : String(out.cooperation_status).trim();

  if(!out.report_date || !/^\d{4}-\d{2}-\d{2}$/.test(out.report_date)){
    return null;
  }

  return out;
}

function parseWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, {
    type:"array",
    cellDates:true
  });

  const records = [];
  const skipped = [];

  for(const sheetName of wb.SheetNames){
    const ws = wb.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(ws, {
      header:1,
      defval:null,
      raw:true
    });

    const h = headerIndex(rows);

    if(h.score < 8){
      skipped.push(sheetName);
      continue;
    }

    let lastAgent = "";

    for(let i = h.row + 1; i < rows.length; i++){
      const row = rows[i] || [];

      const rawAgent = h.map.agent_id === undefined
        ? ""
        : String(row[h.map.agent_id] ?? "").trim();

      if(rawAgent && normalize(rawAgent) !== "总计" && normalize(rawAgent) !== "tổng"){
        lastAgent = rawAgent;
      }

      const rec = rowToRecord(row,h.map,lastAgent);

      if(rec){
        records.push(rec);
      }
    }
  }

  if(!records.length){
    throw new Error("Không tìm thấy dữ liệu hợp lệ trong các sheet. Kiểm tra cột mã đại lý và ngày.");
  }

  const unique = new Map();

  for(const r of records){
    unique.set(`${r.agent_id}__${r.report_date}`,r);
  }

  if(skipped.length){
    console.info("Đã bỏ qua sheet không phải dữ liệu:", skipped);
  }

  return [...unique.values()];
}

// ======================================================
// 5. NHẬP EXCEL VÀO SUPABASE
// ======================================================

async function importExcel(file){
  if(!db){
    throw new Error("Chưa cấu hình Supabase.");
  }

  const records = parseWorkbook(await file.arrayBuffer());

  if(!records.length){
    throw new Error("Không tìm thấy dữ liệu hợp lệ.");
  }

  let count = 0;

  for(let i = 0; i < records.length; i += 250){
    const batch = records.slice(i,i+250);

    const {error} = await db
      .from("agent_stats")
      .upsert(batch, {
        onConflict:"agent_id,report_date"
      });

    if(error) throw error;

    count += batch.length;

    msg(
      "dataMessage",
      `Đang nhập Excel: ${count.toLocaleString("vi-VN")} / ${records.length.toLocaleString("vi-VN")} dòng`
    );
  }

  await loadData();

  msg(
    "dataMessage",
    `Đã nhập/cập nhật ${count.toLocaleString("vi-VN")} dòng từ ${file.name}.`
  );
}

// ======================================================
// 6. QUAN HỆ CẤP TRÊN
// ======================================================

function parentMap(){
  const map = new Map();

  for(const r of hierarchyRows){
    if(r.agent_id && r.parent_agent_id){
      map.set(String(r.agent_id),String(r.parent_agent_id));
    }
  }

  return map;
}

function descendantIds(rootId){
  const root = cleanId(rootId);
  const found = new Set([root]);

  let changed = true;

  while(changed){
    changed = false;

    for(const r of hierarchyRows){
      const child = cleanId(r.agent_id);
      const parent = cleanId(r.parent_agent_id);

      if(parent && found.has(parent) && child && !found.has(child)){
        found.add(child);
        changed = true;
      }
    }
  }

  return found;
}

// Tải quan hệ theo từng trang
async function loadHierarchy(){
  if(!db) return;

  try{
    let allRows = [];

    for(let from = 0; ; from += FETCH_SIZE){
      const {data,error} = await db
        .from("agent_hierarchy")
        .select("agent_id,parent_agent_id,updated_at")
        .order("agent_id", {ascending:true})
        .range(from,from + FETCH_SIZE - 1);

      if(error) throw error;

      allRows = allRows.concat(data || []);

      if(!data || data.length < FETCH_SIZE){
        break;
      }
    }

    hierarchyRows = allRows;

    renderHierarchy();
    render();

  }catch(error){
    msg(
      "hierarchyMessage",
      `Không tải được quan hệ cấp trên: ${error.message}`,
      true
    );

    hierarchyRows = [];
    renderHierarchy();
  }
}

// ======================================================
// 7. HIỂN THỊ DANH SÁCH QUAN HỆ
// ======================================================

function renderHierarchy(){
  renderAgentLinks();

  const body = $("hierarchyBody");

  if(!body) return;

  body.innerHTML = "";

  const rows = [...hierarchyRows].sort((a,b) =>
    String(a.agent_id).localeCompare(
      String(b.agent_id),
      undefined,
      {numeric:true}
    )
  );

  if(!rows.length){
    body.innerHTML =
      '<tr><td colspan="3" class="empty">Chưa có liên kết nào được lưu.</td></tr>';
    return;
  }

  for(const r of rows){
    const tr = document.createElement("tr");

    const child = document.createElement("td");
    child.textContent = r.agent_id;

    const parent = document.createElement("td");
    parent.textContent = r.parent_agent_id || "—";

    const action = document.createElement("td");
    action.style.textAlign = "center";

    const del = document.createElement("button");
    del.className = "danger small-button";
    del.textContent = "Xóa liên kết";

    del.addEventListener("click", () => deleteParent(r.agent_id));

    action.append(del);
    tr.append(child,parent,action);
    body.appendChild(tr);
  }
}

// Hiển thị danh sách mã đại lý từ dữ liệu thống kê
function renderAgentLinks(){
  const body = $("unassignedBody");

  if(!body) return;

  const counts = new Map();

  for(const r of currentRows){
    const id = cleanId(r.agent_id);

    if(id){
      counts.set(id,(counts.get(id) || 0) + 1);
    }
  }

  const parents = parentMap();

  const all = [...counts.entries()].sort((a,b) =>
    a[0].localeCompare(b[0],undefined,{numeric:true})
  );

  const query = normalize($("unassignedSearch")?.value || "");

  const shown = all.filter(([id]) =>
    !query ||
    normalize(id).includes(query) ||
    normalize(parents.get(id) || "").includes(query)
  );

  if($("unassignedCount")){
    $("unassignedCount").textContent =
      `${all.length.toLocaleString("vi-VN")} mã`;
  }

  body.innerHTML = "";

  if(!all.length){
    body.innerHTML =
      '<tr><td colspan="4" class="empty">Chưa có dữ liệu thống kê. Hãy nhập Excel trước.</td></tr>';
    return;
  }

  if(!shown.length){
    body.innerHTML =
      '<tr><td colspan="4" class="empty">Không tìm thấy mã phù hợp.</td></tr>';
    return;
  }

  for(const [id,count] of shown){
    const tr = document.createElement("tr");

    const idCell = document.createElement("td");
    idCell.textContent = id;

    const countCell = document.createElement("td");
    countCell.textContent = count.toLocaleString("vi-VN");

    const parentCell = document.createElement("td");

    const parentInput = document.createElement("input");
    parentInput.type = "text";
    parentInput.className = "inline-parent-input";
    parentInput.placeholder = "Nhập mã cấp trên";
    parentInput.value = parents.get(id) || "";
    parentInput.setAttribute(
      "aria-label",
      `Mã cấp trên trực tiếp của ${id}`
    );

    parentCell.appendChild(parentInput);

    const action = document.createElement("td");
    action.className = "row-actions";

    const link = document.createElement("button");
    link.className = "small-button";
    link.textContent = "Liên kết";

    link.addEventListener("click", () => {
      saveParentFor(id,parentInput.value);
    });

    const unlink = document.createElement("button");
    unlink.className = "danger small-button";
    unlink.textContent = "Xóa liên kết";
    unlink.disabled = !parents.has(id);

    unlink.addEventListener("click", () => {
      deleteParent(id);
    });

    action.append(link,unlink);
    tr.append(idCell,countCell,parentCell,action);

    body.appendChild(tr);
  }
}

// Lưu quan hệ cấp trên
async function saveParentFor(child,parent){
  child = cleanId(child);
  parent = cleanId(parent);

  if(!child){
    msg("hierarchyMessage","Mã đại lý không hợp lệ.",true);
    return;
  }

  if(!parent){
    msg("hierarchyMessage",`Hãy nhập mã cấp trên cho đại lý ${child}.`,true);
    return;
  }

  if(parent === child){
    msg("hierarchyMessage","Một đại lý không thể là cấp trên của chính nó.",true);
    return;
  }

  const map = parentMap();
  let cursor = parent;
  const seen = new Set([child]);

  while(cursor){
    if(seen.has(cursor)){
      msg("hierarchyMessage","Liên kết này tạo vòng lặp cấp trên. Hãy kiểm tra lại.",true);
      return;
    }

    seen.add(cursor);
    cursor = map.get(cursor) || "";
  }

  const {error} = await db
    .from("agent_hierarchy")
    .upsert({
      agent_id:child,
      parent_agent_id:parent,
      updated_at:new Date().toISOString()
    }, {
      onConflict:"agent_id"
    });

  if(error){
    msg("hierarchyMessage",`Không lưu được liên kết: ${error.message}`,true);
    return;
  }

  msg("hierarchyMessage",`Đã liên kết ${child} trực thuộc ${parent}.`);

  await loadHierarchy();
}

// Xóa quan hệ cấp trên
async function deleteParent(child){
  if(!confirm(`Bỏ quan hệ cấp trên của đại lý ${child}?`)){
    return;
  }

  const {error} = await db
    .from("agent_hierarchy")
    .delete()
    .eq("agent_id",child);

  if(error){
    msg("hierarchyMessage",`Không xóa được quan hệ: ${error.message}`,true);
    return;
  }

  msg("hierarchyMessage",`Đã bỏ quan hệ cấp trên của ${child}.`);

  await loadHierarchy();
}

// ======================================================
// 8. CHUYỂN TRANG
// ======================================================

function showView(view){
  $("statsView").classList.toggle("hidden",view !== "stats");
  $("hierarchyView").classList.toggle("hidden",view !== "hierarchy");

  $("statsTab").classList.toggle("active",view === "stats");
  $("hierarchyTab").classList.toggle("active",view === "hierarchy");
}

// ======================================================
// 9. TẢI TOÀN BỘ DỮ LIỆU THEO TỪNG ĐỢT
// ======================================================

async function loadData(){
  if(!db || loadingData) return;

  loadingData = true;

  msg("dataMessage","Đang tải dữ liệu...");

  try{
    let allRows = [];

    for(let from = 0; ; from += FETCH_SIZE){
      const {data,error} = await db
        .from("agent_stats")
        .select("*")
        .order("id",{ascending:true})
        .range(from,from + FETCH_SIZE - 1);

      if(error) throw error;

      allRows = allRows.concat(data || []);

      msg(
        "dataMessage",
        `Đang tải dữ liệu: ${allRows.length.toLocaleString("vi-VN")} dòng`
      );

      if(!data || data.length < FETCH_SIZE){
        break;
      }
    }

    currentRows = allRows;
    statsPage = 1;

    render();

    await loadHierarchy();

    msg(
      "dataMessage",
      `Đã tải thành công ${currentRows.length.toLocaleString("vi-VN")} dòng dữ liệu.`
    );

  }catch(error){
    msg(
      "dataMessage",
      `Tải dữ liệu thất bại: ${error.message || error}`,
      true
    );

  }finally{
    loadingData = false;
  }
}

// ======================================================
// 10. LỌC DỮ LIỆU
// ======================================================

function filteredRows(){
  const rawAgent = cleanId($("agentFilter").value);
  const agent = normalize(rawAgent);

  const from = $("dateFrom").value;
  const to = $("dateTo").value;

  let allowed = null;

  if(rawAgent){
    const exactExists =
      currentRows.some(r => normalize(r.agent_id) === agent) ||
      hierarchyRows.some(r =>
        normalize(r.agent_id) === agent ||
        normalize(r.parent_agent_id) === agent
      );

    if(exactExists){
      allowed = descendantIds(rawAgent);
    }
  }

  return currentRows.filter(r => {
    const id = String(r.agent_id ?? "");

    const matchesAgent =
      !agent ||
      (allowed
        ? allowed.has(id)
        : normalize(id).includes(agent));

    return matchesAgent &&
      (!from || r.report_date >= from) &&
      (!to || r.report_date <= to);
  });
}

// ======================================================
// 11. HIỂN THỊ THỐNG KÊ THEO TỪNG TRANG
// ======================================================

function render(){
  const rows = filteredRows();

  const body = $("statsBody");
  const foot = $("statsFoot");

  // Tính tổng trên toàn bộ dữ liệu sau khi lọc
  $("kpiAgents").textContent =
    new Set(rows.map(r => r.agent_id)).size.toLocaleString("vi-VN");

  $("kpiRows").textContent =
    rows.length.toLocaleString("vi-VN");

  const sum = f =>
    rows.reduce((a,r) => a + (Number(r[f]) || 0),0);

  $("kpiExpense").textContent = fmt(sum("total_expense"));
  $("kpiProfit").textContent = fmt(sum("total_profit"));

  body.innerHTML = "";
  foot.innerHTML = "";

  const totalPages = Math.max(
    1,
    Math.ceil(rows.length / PAGE_SIZE)
  );

  if(statsPage > totalPages){
    statsPage = totalPages;
  }

  if(statsPage < 1){
    statsPage = 1;
  }

  if(!rows.length){
    body.innerHTML =
      '<tr><td colspan="20" class="empty">Không có dữ liệu phù hợp bộ lọc.</td></tr>';

    updateStatsPagination(0,totalPages);
    return;
  }

  // Chỉ lấy 200 dòng của trang hiện tại
  const start = (statsPage - 1) * PAGE_SIZE;

  const pageRows = rows.slice(
    start,
    start + PAGE_SIZE
  );

  const displayFields = COLS.map(c => c[0]);

  for(const r of pageRows){
    const tr = document.createElement("tr");

    for(const f of displayFields){
      const td = document.createElement("td");

      let v = r[f];

      if(f === "report_date"){
        v = dateLabel(v);
      }else if(NUMERIC_FIELDS.includes(f)){
        v = fmt(v);
      }else{
        v = v ?? "";
      }

      td.textContent = v;

      if(
        NUMERIC_FIELDS.includes(f) &&
        typeof r[f] === "number" &&
        r[f] < 0
      ){
        td.classList.add("negative");
      }

      tr.appendChild(td);
    }

    body.appendChild(tr);
  }

  // Dòng tổng tính trên tất cả dữ liệu đã lọc
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

  const tr = document.createElement("tr");

  for(const [i,f] of displayFields.entries()){
    const td = document.createElement("td");

    if(i === 0){
      td.textContent = "TỔNG";
    }else if(f === "report_date"){
      td.textContent = rows.length + " dòng";
    }else if(totalFields.has(f)){
      const n = sum(f);

      td.textContent = fmt(n);

      if(n < 0){
        td.classList.add("negative");
      }
    }else if(f === "profit_percent"){
      const denominator = sum("total_expense");

      td.textContent = denominator
        ? fmt(sum("total_profit") / denominator * 100) + "%"
        : "—";
    }else{
      td.textContent = "—";
    }

    tr.appendChild(td);
  }

  foot.appendChild(tr);

  updateStatsPagination(rows.length,totalPages);
}

// ======================================================
// 12. TẠO NÚT PHÂN TRANG
// ======================================================

function updateStatsPagination(totalRows,totalPages){
  let box = $("statsPagination");

  if(!box){
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

    $("statsTable").parentElement.insertAdjacentElement(
      "afterend",
      box
    );
  }

  const start = totalRows
    ? (statsPage - 1) * PAGE_SIZE + 1
    : 0;

  const end = Math.min(
    statsPage * PAGE_SIZE,
    totalRows
  );

  box.innerHTML = `
    <button id="statsPrev" class="secondary" type="button"
      ${statsPage <= 1 ? "disabled" : ""}>
      Trang trước
    </button>

    <span>
      Trang ${statsPage} / ${totalPages}
      · Hiển thị ${start}-${end}
      trên ${totalRows.toLocaleString("vi-VN")} dòng
    </span>

    <button id="statsNext" class="secondary" type="button"
      ${statsPage >= totalPages ? "disabled" : ""}>
      Trang sau
    </button>
  `;

  $("statsPrev").addEventListener("click",() => {
    if(statsPage > 1){
      statsPage--;
      render();
    }
  });

  $("statsNext").addEventListener("click",() => {
    if(statsPage < totalPages){
      statsPage++;
      render();
    }
  });
}

// ======================================================
// 13. ĐĂNG NHẬP VÀ KHỞI ĐỘNG
// ======================================================

function setSignedIn(yes){
  $("loginPanel").classList.toggle("hidden",yes);
  $("appPanel").classList.toggle("hidden",!yes);
  $("logoutBtn").classList.toggle("hidden",!yes);

  $("accountStatus").textContent =
    yes ? "Đã đăng nhập" : "Chưa đăng nhập";
}

async function init(){
  if(!configured()){
    $("setupNotice").classList.remove("hidden");
    $("loginPanel").classList.remove("hidden");

    $("loginMessage").textContent =
      "Hãy điền thông tin Supabase trong app.js trước khi sử dụng.";

    $("loginForm").querySelectorAll("input,button").forEach(el => {
      el.disabled = true;
    });

    return;
  }

  db = window.supabase.createClient(
    SUPABASE_URL,
    SUPABASE_ANON_KEY
  );

  $("setupNotice").classList.add("hidden");

  // Chỉ cập nhật giao diện đăng nhập tại đây.
  // Không gọi loadData trong callback để tránh tải lặp.
  db.auth.onAuthStateChange((_event,session) => {
    setSignedIn(!!session);
  });

  const {data:{session}} = await db.auth.getSession();

  setSignedIn(!!session);

  if(session){
    await loadData();
  }
}

// ======================================================
// 14. SỰ KIỆN ĐĂNG NHẬP
// ======================================================

$("loginForm").addEventListener("submit",async e => {
  e.preventDefault();

  if(!db) return;

  msg("loginMessage","Đang đăng nhập...");

  const {error} = await db.auth.signInWithPassword({
    email:$("email").value.trim(),
    password:$("password").value
  });

  msg(
    "loginMessage",
    error ? error.message : "Đăng nhập thành công.",
    !!error
  );

  if(!error){
    await loadData();
  }
});

// ======================================================
// 15. TẠO TÀI KHOẢN
// ======================================================

$("signupBtn").addEventListener("click",async() => {
  if(!db) return;

  msg("loginMessage","Đang tạo tài khoản...");

  const {data,error} = await db.auth.signUp({
    email:$("email").value.trim(),
    password:$("password").value
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

  if(!error && data.session){
    await loadData();
  }
});

// ======================================================
// 16. ĐĂNG XUẤT
// ======================================================

$("logoutBtn").addEventListener("click",async() => {
  if(db){
    await db.auth.signOut();
  }
});

// ======================================================
// 17. NHẬP FILE EXCEL
// ======================================================

$("excelFile").addEventListener("change",async e => {
  const file = e.target.files?.[0];

  if(!file) return;

  try{
    msg("dataMessage","Đang đọc file Excel...");

    await importExcel(file);

  }catch(err){
    msg(
      "dataMessage",
      `Nhập Excel thất bại: ${err.message || err}`,
      true
    );

  }finally{
    e.target.value = "";
  }
});

// ======================================================
// 18. BỘ LỌC
// ======================================================

["agentFilter","dateFrom","dateTo"].forEach(id => {
  $(id).addEventListener("input",() => {
    statsPage = 1;
    render();
  });
});

$("clearFiltersBtn").addEventListener("click",() => {
  $("agentFilter").value = "";
  $("dateFrom").value = "";
  $("dateTo").value = "";

  statsPage = 1;

  render();
});

// ======================================================
// 19. CÁC NÚT TẢI LẠI VÀ CHUYỂN TRANG
// ======================================================

$("refreshBtn").addEventListener("click",loadData);

$("statsTab").addEventListener("click",() => {
  showView("stats");
});

$("hierarchyTab").addEventListener("click",() => {
  showView("hierarchy");
});

$("refreshHierarchyBtn").addEventListener("click",loadHierarchy);

$("unassignedSearch").addEventListener("input",renderAgentLinks);

// ======================================================
// 20. XUẤT EXCEL
// ======================================================

$("exportBtn").addEventListener("click",() => {
  const rows = filteredRows();

  if(!rows.length){
    msg("dataMessage","Không có dữ liệu để xuất.",true);
    return;
  }

  const aoa = [
    COLS.map(c => c[1][0])
  ];

  rows.forEach(r => {
    aoa.push(
      COLS.map(([f]) =>
        f === "report_date"
          ? dateLabel(r[f])
          : r[f] ?? ""
      )
    );
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    "ThongKe"
  );

  XLSX.writeFile(
    wb,
    "thong-ke-dai-ly.xlsx"
  );
});

// ======================================================
// 21. CHẠY WEBSITE
// ======================================================

init();
