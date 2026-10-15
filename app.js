/* Chấm Công Workspace — giao diện mẫu v1.
   Dữ liệu bên dưới chỉ là dữ liệu demo trong bộ nhớ, chưa kết nối cơ sở dữ liệu.
   Khi phát triển tiếp, thay phần state bằng Supabase/API và giữ nguyên giao diện. */
(() => {
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

  const demoRows = [
    {id:1,code:"NV001",name:"Nguyễn Văn An",department:"Sản xuất",checkIn:"07:55",checkOut:"17:05",status:"Có mặt",color:"blue"},
    {id:2,code:"NV002",name:"Trần Thị Bình",department:"Hành chính",checkIn:"08:10",checkOut:"17:00",status:"Đi muộn",color:"purple"},
    {id:3,code:"NV003",name:"Lê Minh Cường",department:"Sản xuất",checkIn:"07:50",checkOut:"17:10",status:"Có mặt",color:"green"},
    {id:4,code:"NV004",name:"Phạm Thị Dung",department:"Kế toán",checkIn:"—",checkOut:"—",status:"Vắng mặt",color:"orange"},
    {id:5,code:"NV005",name:"Hoàng Văn Hải",department:"Kho",checkIn:"08:05",checkOut:"17:00",status:"Có mặt",color:"blue"},
    {id:6,code:"NV006",name:"Đặng Thu Hà",department:"Sản xuất",checkIn:"—",checkOut:"—",status:"Vắng mặt",color:"pink"},
    {id:7,code:"NV007",name:"Bùi Minh Khôi",department:"Kỹ thuật",checkIn:"07:45",checkOut:"17:15",status:"Có mặt",color:"green"},
    {id:8,code:"NV008",name:"Ngô Thị Lan",department:"Hành chính",checkIn:"08:12",checkOut:"—",status:"Đi muộn",color:"purple"}
  ];
  let rows = demoRows.map(row => ({...row}));
  let activeFilter = "all";
  let editingId = null;
  let searchTerm = "";

  const escapeHTML = value => String(value ?? "").replace(/[&<>"']/g, char => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[char]));
  const initials = name => name.trim().split(/\s+/).slice(-2).map(part => part[0]).join("").toUpperCase();
  const statusClass = status => status === "Có mặt" ? "present" : status === "Đi muộn" ? "late" : "absent";

  function setDateLabels() {
    const now = new Date();
    $("#fullDate").textContent = new Intl.DateTimeFormat("vi-VN", {weekday:"long",day:"2-digit",month:"2-digit",year:"numeric"}).format(now);
    $("#todayLabel").textContent = new Intl.DateTimeFormat("vi-VN", {day:"2-digit",month:"2-digit",year:"numeric"}).format(now);
    $("#dateFilter").value = [now.getFullYear(),String(now.getMonth()+1).padStart(2,"0"),String(now.getDate()).padStart(2,"0")].join("-");
  }

  function renderRows() {
    const filtered = rows.filter(row => {
      const matchesFilter = activeFilter === "all" || row.status === activeFilter;
      const haystack = `${row.name} ${row.code} ${row.department}`.toLocaleLowerCase("vi");
      return matchesFilter && haystack.includes(searchTerm.toLocaleLowerCase("vi"));
    });
    $("#attendanceRows").innerHTML = filtered.map(row => `
      <tr>
        <td><div class="employee-cell"><div class="employee-avatar ${escapeHTML(row.color)}">${escapeHTML(initials(row.name))}</div><div><div class="employee-name">${escapeHTML(row.name)}</div><div class="employee-code">${escapeHTML(row.code)}</div></div></div></td>
        <td><span class="department">${escapeHTML(row.department || "—")}</span></td>
        <td>${escapeHTML(row.checkIn || "—")}</td><td>${escapeHTML(row.checkOut || "—")}</td>
        <td><span class="status ${statusClass(row.status)}">${escapeHTML(row.status)}</span></td>
        <td><div class="row-actions"><button class="row-action" data-action="edit" data-id="${row.id}" aria-label="Sửa ${escapeHTML(row.name)}" title="Sửa">✎</button><button class="row-action delete" data-action="delete" data-id="${row.id}" aria-label="Xóa ${escapeHTML(row.name)}" title="Xóa">×</button></div></td>
      </tr>`).join("");
    $("#emptyState").classList.toggle("hidden", filtered.length > 0);
    $("#resultCount").textContent = `Hiển thị ${filtered.length} bản ghi mẫu`;
    $("#allTabCount").textContent = rows.length;
    updateStats();
  }

  function updateStats() {
    const present = rows.filter(row => row.status === "Có mặt").length;
    const late = rows.filter(row => row.status === "Đi muộn").length;
    const absent = rows.filter(row => row.status === "Vắng mặt").length;
    $("#presentCount").textContent = present;
    $("#lateCount").textContent = late;
    $("#absentCount").textContent = absent;
    $("#legendPresent").textContent = present;
    $("#legendLate").textContent = late;
    $("#legendAbsent").textContent = absent;
    const total = rows.length;
    $("#totalEmployees").textContent = Math.max(25, total);
    $("#attendanceRate").textContent = `${total ? Math.round((present / total) * 100) : 0}%`;
    const green = total ? (present / total) * 100 : 0;
    const orange = total ? (late / total) * 100 : 0;
    $(".donut").style.background = `conic-gradient(#20b483 0 ${green}%,#f1b64e ${green}% ${green+orange}%,#ed6971 ${green+orange}% 100%)`;
  }

  function showToast(title, message) {
    $("#toastTitle").textContent = title;
    $("#toastMessage").textContent = message;
    $("#toast").classList.remove("hidden");
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => $("#toast").classList.add("hidden"), 2800);
  }
  function openModal(row = null) {
    editingId = row ? row.id : null;
    const form = $("#attendanceForm");
    form.reset();
    $("#modalTitle").textContent = row ? "Sửa thông tin chấm công" : "Thêm chấm công";
    if (row) {
      form.elements.name.value = row.name;
      form.elements.code.value = row.code;
      form.elements.department.value = row.department;
      form.elements.checkIn.value = row.checkIn === "—" ? "" : row.checkIn;
      form.elements.checkOut.value = row.checkOut === "—" ? "" : row.checkOut;
      form.elements.status.value = row.status;
    }
    $("#modalOverlay").classList.remove("hidden");
    setTimeout(() => form.elements.name.focus(), 40);
  }
  function closeModal() {
    $("#modalOverlay").classList.add("hidden");
    editingId = null;
  }

  $("#attendanceRows").addEventListener("click", event => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const id = Number(button.dataset.id);
    const row = rows.find(item => item.id === id);
    if (!row) return;
    if (button.dataset.action === "edit") openModal(row);
    if (button.dataset.action === "delete" && confirm(`Xóa bản ghi mẫu của ${row.name}?`)) {
      rows = rows.filter(item => item.id !== id);
      renderRows();
      showToast("Đã xóa bản ghi mẫu", "Thay đổi chỉ tồn tại trong phiên xem trước.");
    }
  });

  $("#attendanceForm").addEventListener("submit", event => {
    event.preventDefault();
    const form = event.currentTarget;
    const name = form.elements.name.value.trim();
    const code = form.elements.code.value.trim().toUpperCase();
    if (rows.some(row => row.code.toLowerCase() === code.toLowerCase() && row.id !== editingId)) {
      showToast("Mã nhân viên đã tồn tại", "Vui lòng sử dụng một mã khác.");
      form.elements.code.focus();
      return;
    }
    const data = {
      id: editingId ?? (Math.max(0, ...rows.map(row => row.id)) + 1),
      name, code, department: form.elements.department.value.trim(),
      checkIn: form.elements.checkIn.value || "—",
      checkOut: form.elements.checkOut.value || "—",
      status: form.elements.status.value,
      color: ["blue","green","purple","orange","pink"][Math.floor(Math.random()*5)]
    };
    if (editingId !== null) rows = rows.map(row => row.id === editingId ? data : row);
    else rows.unshift(data);
    closeModal();
    renderRows();
    showToast(editingId === null ? "Đã thêm bản ghi mẫu" : "Đã cập nhật bản ghi mẫu", "Lưu ý: dữ liệu chưa được lưu lên máy chủ.");
  });

  $("#addAttendanceBtn").addEventListener("click", () => openModal());
  $("#closeModal").addEventListener("click", closeModal);
  $("#cancelModal").addEventListener("click", closeModal);
  $("#modalOverlay").addEventListener("click", event => { if (event.target === $("#modalOverlay")) closeModal(); });
  document.addEventListener("keydown", event => { if (event.key === "Escape") closeModal(); });
  $("#globalSearch").addEventListener("input", event => { searchTerm = event.target.value.trim(); renderRows(); });

  $$(".filter-tab").forEach(button => button.addEventListener("click", () => {
    activeFilter = button.dataset.filter;
    $$(".filter-tab").forEach(tab => tab.classList.toggle("selected", tab === button));
    renderRows();
  }));

  const pages = {
    dashboard: ["Tổng quan", "Chào bạn, Admin ✦", "Theo dõi tình hình chấm công và nhân sự của bạn tại một nơi."],
    attendance: ["Chấm công", "Quản lý chấm công", "Theo dõi giờ vào, giờ ra và trạng thái của từng nhân viên."],
    employees: ["Nhân viên", "Quản lý nhân viên", "Danh sách nhân sự minh họa — có thể tìm kiếm và chỉnh sửa bản ghi."],
    reports: ["Thống kê & báo cáo", "Thống kê & báo cáo", "Các chỉ số đang hiển thị là dữ liệu mẫu, chưa phải báo cáo thực tế."],
    settings: ["Cài đặt", "Cài đặt hệ thống", "Khu vực dành cho các tùy chỉnh sẽ được bổ sung trong những phiên bản sau."]
  };
  function navigate(page) {
    const details = pages[page] || pages.dashboard;
    $("#breadcrumbTitle").textContent = details[0];
    $("#pageTitle").textContent = details[1];
    $("#pageSubtitle").textContent = details[2];
    $$(".nav-link").forEach(button => button.classList.toggle("active", button.dataset.page === page));
    $("#sidebar").classList.remove("open");
    if (page === "attendance") $("#attendancePanel").scrollIntoView({behavior:"smooth",block:"start"});
    else if (page === "employees") $("#attendancePanel").scrollIntoView({behavior:"smooth",block:"start"});
    else if (page === "reports") $(".summary-panel").scrollIntoView({behavior:"smooth",block:"center"});
    else if (page === "settings") showToast("Cài đặt", "Mục này sẽ được bổ sung ở phiên bản sau.");
    else window.scrollTo({top:0,behavior:"smooth"});
  }
  $$("[data-page]").forEach(button => button.addEventListener("click", () => navigate(button.dataset.page)));
  $("#menuToggle").addEventListener("click", () => $("#sidebar").classList.toggle("open"));
  $("#learnMore").addEventListener("click", () => showToast("Bản giao diện 1.0", "Các bản ghi là dữ liệu demo, chưa lưu lên máy chủ."));
  $("#dateFilter").addEventListener("change", () => showToast("Đã chọn ngày", "Bộ lọc ngày sẽ được kết nối với dữ liệu thật ở phiên bản sau."));

  $("#exportBtn").addEventListener("click", () => {
    const filtered = rows.filter(row => {
      const matchesFilter = activeFilter === "all" || row.status === activeFilter;
      const haystack = `${row.name} ${row.code} ${row.department}`.toLocaleLowerCase("vi");
      return matchesFilter && haystack.includes(searchTerm.toLocaleLowerCase("vi"));
    });
    const csvCell = value => `"${String(value ?? "").replace(/"/g,'""')}"`;
    const csv = ["Mã nhân viên,Tên nhân viên,Phòng ban,Giờ vào,Giờ ra,Trạng thái",
      ...filtered.map(row => [row.code,row.name,row.department,row.checkIn,row.checkOut,row.status].map(csvCell).join(","))].join("\r\n");
    const blob = new Blob(["\ufeff", csv], {type:"text/csv;charset=utf-8;"});
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "du-lieu-cham-cong-mau.csv";
    document.body.appendChild(link); link.click(); link.remove();
    URL.revokeObjectURL(url);
    showToast("Đã xuất CSV", "Tệp chứa dữ liệu mẫu đang được tải xuống.");
  });

  setDateLabels();
  renderRows();
})();