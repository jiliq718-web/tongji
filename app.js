// 1) Điền thông tin Supabase tại đây (Project Settings > API).
const SUPABASE_URL = "https://zuzyekphhcwwidcbvnqq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_OxHoqPnBjsvYos7WUm2kpg_hk8GEzyd";

// Tên cột nội bộ và các tiêu đề Excel thường gặp (Việt/Trung).
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
  ["profit_percent", ["利润百分比","利润百分比","LN %","Tỷ lệ lợi nhuận"]],
  ["handling_fee", ["手续费","Phí thủ tục"]],
  ["new_ad_status", ["新广告费是否异常","Quảng cáo mới","Tình trạng quảng cáo"]],
  ["cooperation_status", ["合作状态","Trạng thái hợp tác"]]
];
const NUMERIC_FIELDS = COLS.slice(2,18).map(x=>x[0]);
let db = null, currentRows = [];

const $ = id => document.getElementById(id);
function configured(){return SUPABASE_URL.startsWith("https://") && !SUPABASE_URL.includes("PASTE_") && !SUPABASE_ANON_KEY.includes("PASTE_") && SUPABASE_ANON_KEY.length>20}
function msg(id,text,error=false){const el=$(id);el.textContent=text;el.classList.toggle("error",error)}
function normalize(s){return String(s??"").trim().toLowerCase().replace(/\s+/g,"").replace(/[：:]/g,"")}
function parseNumber(v){
  if(typeof v==="number") return Number.isFinite(v)?v:null;
  if(v===null||v===undefined||String(v).trim()==="") return null;
  const s=String(v).replace(/[,，￥¥%\s]/g,"").replace(/[()]/g,"");
  const n=Number(s); return Number.isFinite(n)?n:null;
}
function parseDate(v){
  if(v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  if(typeof v==="number" && v>20000 && v<80000){
    const d=XLSX.SSF.parse_date_code(v); if(d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
  }
  const s=String(v??"").trim();
  let m=s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);
  if(m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  m=s.match(/^(\d{1,2})月(\d{1,2})(?:日|号)?$/);
  if(m) return `${new Date().getFullYear()}-${m[1].padStart(2,"0")}-${m[2].padStart(2,"0")}`;
  m=s.match(/^(\d{1,2})[\/.-](\d{1,2})$/);
  if(m) return `${new Date().getFullYear()}-${m[1].padStart(2,"0")}-${m[2].padStart(2,"0")}`;
  return s;
}
function dateLabel(s){if(!s)return "";const d=new Date(`${s}T00:00:00`);return isNaN(d)?s:`${d.getMonth()+1}月${d.getDate()}号`}
function fmt(v){if(v===null||v===undefined||v==="")return "—";if(typeof v==="number")return v.toLocaleString("vi-VN",{maximumFractionDigits:4});return String(v)}
function headerIndex(rows){
  let best={row:-1,score:0,map:{}};
  for(let r=0;r<Math.min(rows.length,20);r++){
    const cells=(rows[r]||[]).map(normalize), map={}; let score=0;
    for(const [key,aliases] of COLS){
      const idx=cells.findIndex(c=>[key,...aliases.map(normalize)].includes(c));
      if(idx>=0){map[key]=idx;score++}
    }
    if(score>best.score)best={row:r,score,map};
  }
  return best;
}
function rowToRecord(row,map,lastAgent){
  const out={}; for(const [key] of COLS){const idx=map[key];out[key]=idx===undefined?null:row[idx]}
  let agent=String(out.agent_id??"").trim();
  if(!agent || agent==="undefined" || agent==="null") agent=lastAgent||"";
  if(!agent || normalize(agent)==="总计" || normalize(agent)==="tổng") return null;
  out.agent_id=agent;
  out.report_date=parseDate(out.report_date);
  for(const f of NUMERIC_FIELDS){
    if(f==="profit_percent" && typeof out[f]==="string" && out[f].includes("%")){
      const n=parseNumber(out[f]);out[f]=n===null?null:n;
    }else out[f]=parseNumber(out[f]);
  }
  out.new_ad_status=out.new_ad_status==null?"":String(out.new_ad_status).trim();
  out.cooperation_status=out.cooperation_status==null?"":String(out.cooperation_status).trim();
  if(!out.report_date || !/^\d{4}-\d{2}-\d{2}$/.test(out.report_date)) return null;
  return out;
}
function parseWorkbook(arrayBuffer){
  const wb=XLSX.read(arrayBuffer,{type:"array",cellDates:true});
  const records=[];
  const skipped=[];
  // File WPS của bạn có nhiều sheet đại lý; đọc tất cả sheet có hàng tiêu đề phù hợp.
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    const h=headerIndex(rows);
    // Bỏ qua sheet phụ/trống thay vì làm hỏng toàn bộ lần nhập.
    if(h.score<8){skipped.push(sheetName);continue;}
    let lastAgent="";
    for(let i=h.row+1;i<rows.length;i++){
      const row=rows[i]||[];
      const rawAgent=h.map.agent_id===undefined?"":String(row[h.map.agent_id]??"").trim();
      if(rawAgent && normalize(rawAgent)!=="总计" && normalize(rawAgent)!=="tổng")lastAgent=rawAgent;
      const rec=rowToRecord(row,h.map,lastAgent);
      if(rec)records.push(rec);
    }
  }
  if(!records.length) throw new Error("Không tìm thấy dữ liệu hợp lệ trong các sheet. File cần có tiêu đề 代理ID, 日期 và các cột số liệu.");
  // Tránh gửi nhiều dòng trùng khóa (mã đại lý + ngày) trong cùng một lô nhập.
  const unique=new Map();
  for(const r of records) unique.set(`${r.agent_id}__${r.report_date}`,r);
  if(skipped.length) console.info("Đã bỏ qua sheet không phải dữ liệu:",skipped);
  return [...unique.values()];
}
async function importExcel(file){
  if(!db)throw new Error("Chưa cấu hình Supabase.");
  const records=parseWorkbook(await file.arrayBuffer());
  if(!records.length)throw new Error("Không tìm thấy dòng dữ liệu hợp lệ. Kiểm tra cột ngày và mã đại lý.");
  // Gửi theo từng lô nhỏ, tránh request quá lớn.
  let count=0;
  for(let i=0;i<records.length;i+=250){
    const batch=records.slice(i,i+250);
    const {error}=await db.from("agent_stats").upsert(batch,{onConflict:"agent_id,report_date"});
    if(error)throw error;
    count+=batch.length;
  }
  msg("dataMessage",`Đã nhập/cập nhật ${count} dòng từ ${file.name}.`);
  await loadData();
}
async function loadData(){
  if(!db)return;
  msg("dataMessage","Đang tải dữ liệu...");
  const {data,error}=await db.from("agent_stats").select("*").order("report_date",{ascending:true}).limit(10000);
  if(error){msg("dataMessage",error.message,true);return}
  currentRows=data||[];render();
  msg("dataMessage",`Đã tải ${currentRows.length} dòng dữ liệu.`);
}
function filteredRows(){
  const agent=normalize($("agentFilter").value);
  const from=$("dateFrom").value,to=$("dateTo").value;
  return currentRows.filter(r=>(!agent||normalize(r.agent_id).includes(agent))&&(!from||r.report_date>=from)&&(!to||r.report_date<=to));
}
function render(){
  const rows=filteredRows(),body=$("statsBody"),foot=$("statsFoot");
  $("kpiAgents").textContent=new Set(rows.map(r=>r.agent_id)).size.toLocaleString("vi-VN");
  $("kpiRows").textContent=rows.length.toLocaleString("vi-VN");
  const sum=f=>rows.reduce((a,r)=>a+(Number(r[f])||0),0);
  $("kpiExpense").textContent=fmt(sum("total_expense"));
  $("kpiProfit").textContent=fmt(sum("total_profit"));
  body.innerHTML="";foot.innerHTML="";
  if(!rows.length){body.innerHTML='<tr><td colspan="20" class="empty">Không có dữ liệu phù hợp bộ lọc.</td></tr>';return}
  const displayFields=COLS.map(c=>c[0]);
  for(const r of rows){
    const tr=document.createElement("tr");
    for(const f of displayFields){
      const td=document.createElement("td");
      let v=r[f];
      if(f==="report_date")v=dateLabel(v);
      else if(NUMERIC_FIELDS.includes(f))v=fmt(v);
      else v=v??"";
      td.textContent=v;
      if(NUMERIC_FIELDS.includes(f)&&typeof r[f]==="number"&&r[f]<0)td.classList.add("negative");
      tr.appendChild(td);
    }
    body.appendChild(tr);
  }
  const totalFields=new Set(["daily_salary","ad_funds","commission","video_promotion","total_expense","level1_deposit","level1_withdrawal","level2_deposit","total_assets","total_withdrawals","internal_profit","external_profit","storage_diff","total_profit","handling_fee"]);
  const tr=document.createElement("tr");
  for(const [i,f] of displayFields.entries()){
    const td=document.createElement("td");
    if(i===0)td.textContent="TỔNG";
    else if(f==="report_date")td.textContent=String(rows.length)+" dòng";
    else if(totalFields.has(f)){const n=sum(f);td.textContent=fmt(n);if(n<0)td.classList.add("negative")}
    else if(f==="profit_percent"){const denominator=sum("total_expense");td.textContent=denominator?fmt(sum("total_profit")/denominator*100)+"%":"—"}
    else td.textContent="—";
    tr.appendChild(td);
  }
  foot.appendChild(tr);
}
async function init(){
  if(!configured()){
    $("setupNotice").classList.remove("hidden");
    $("loginPanel").classList.remove("hidden");
    $("loginMessage").textContent="Hãy điền thông tin Supabase trong app.js trước khi sử dụng.";
    $("loginForm").querySelectorAll("input,button").forEach(el=>el.disabled=true);
    return;
  }
  db=window.supabase.createClient(SUPABASE_URL,SUPABASE_ANON_KEY);
  $("setupNotice").classList.add("hidden");
  db.auth.onAuthStateChange((_event,session)=>{setSignedIn(!!session);if(session)loadData()});
  const {data:{session}}=await db.auth.getSession();
  setSignedIn(!!session);
  if(session)await loadData();
}
function setSignedIn(yes){
  $("loginPanel").classList.toggle("hidden",yes);
  $("appPanel").classList.toggle("hidden",!yes);
  $("logoutBtn").classList.toggle("hidden",!yes);
  $("accountStatus").textContent=yes?"Đã đăng nhập":"Chưa đăng nhập";
}
$("loginForm").addEventListener("submit",async e=>{
  e.preventDefault();if(!db)return;
  msg("loginMessage","Đang đăng nhập...");
  const {error}=await db.auth.signInWithPassword({email:$("email").value.trim(),password:$("password").value});
  msg("loginMessage",error?error.message:"Đăng nhập thành công.",!!error);
});
$("signupBtn").addEventListener("click",async()=>{
  if(!db)return;
  msg("loginMessage","Đang tạo tài khoản...");
  const {data,error}=await db.auth.signUp({email:$("email").value.trim(),password:$("password").value});
  msg("loginMessage",error?error.message:(data.session?"Tạo tài khoản thành công.":"Đã gửi email xác nhận nếu dự án yêu cầu xác minh."),!!error);
});
$("logoutBtn").addEventListener("click",async()=>{if(db)await db.auth.signOut()});
$("excelFile").addEventListener("change",async e=>{
  const file=e.target.files?.[0];if(!file)return;
  try{msg("dataMessage","Đang đọc file Excel...");await importExcel(file)}
  catch(err){msg("dataMessage",`Nhập Excel thất bại: ${err.message||err}`,true)}
  finally{e.target.value=""}
});
["agentFilter","dateFrom","dateTo"].forEach(id=>$(id).addEventListener("input",render));
$("clearFiltersBtn").addEventListener("click",()=>{$("agentFilter").value="";$("dateFrom").value="";$("dateTo").value="";render()});
$("refreshBtn").addEventListener("click",loadData);
$("exportBtn").addEventListener("click",()=>{
  const rows=filteredRows();
  if(!rows.length){msg("dataMessage","Không có dữ liệu để xuất.",true);return}
  const aoa=[COLS.map(c=>c[1][0])];
  rows.forEach(r=>aoa.push(COLS.map(([f])=>f==="report_date"?dateLabel(r[f]):r[f]??"")));
  const ws=XLSX.utils.aoa_to_sheet(aoa),wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"ThongKe");
  XLSX.writeFile(wb,"thong-ke-dai-ly.xlsx");
});
init();
