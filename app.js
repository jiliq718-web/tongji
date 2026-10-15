
/* ==========================================
   CHẤM CÔNG WORKSPACE - APP.JS
   Phiên bản 1: giao diện và dữ liệu mẫu
   Chưa kết nối cơ sở dữ liệu dùng chung
   ========================================== */

(() => {
  "use strict";

  const $ = (selector, root = document) =>
    root.querySelector(selector);

  const $$ = (selector, root = document) =>
    [...root.querySelectorAll(selector)];

  // ------------------------------------------
  // 1. DỮ LIỆU MẪU
  // ------------------------------------------

  let rows = [
    {
      id: 1,
      code: "NV001",
      name: "Nguyễn Văn An",
      department: "Sản xuất",
      checkIn: "07:55",
      checkOut: "17:05",
      status: "Có mặt"
    },
    {
      id: 2,
      code: "NV002",
      name: "Trần Thị Bình",
      department: "Hành chính",
      checkIn: "08:10",
      checkOut: "17:00",
      status: "Đi muộn"
    },
    {
      id: 3,
      code: "NV003",
      name: "Lê Minh Cường",
      department: "Sản xuất",
      checkIn: "07:50",
      checkOut: "17:10",
      status: "Có mặt"
    },
    {
      id: 4,
      code: "NV004",
      name: "Phạm Thị Dung",
      department: "Kế toán",
      checkIn: "—",
      checkOut: "—",
      status: "Vắng mặt"
    },
    {
      id: 5,
      code: "NV005",
      name: "Hoàng Văn Hải",
      department: "Kho",
      checkIn: "08:05",
      checkOut: "17:00",
      status: "Có mặt"
    },
    {
      id: 6,
      code: "NV006",
      name: "Đặng Thu Hà",
      department: "Sản xuất",
      checkIn: "—",
      checkOut: "—",
      status: "Vắng mặt"
    },
    {
      id: 7,
      code: "NV007",
      name: "Bùi Minh Khôi",
      department: "Kỹ thuật",
      checkIn: "07:45",
      checkOut: "17:15",
      status: "Có mặt"
    },
    {
      id: 8,
      code: "NV008",
      name: "Ngô Thị Lan",
      department: "Hành chính",
      checkIn: "08:12",
      checkOut: "—",
      status: "Đi muộn"
    }
  ];

  let activeFilter = "all";
  let searchTerm = "";
  let editingId = null;

  // ------------------------------------------
  // 2. CÁC HÀM TIỆN ÍCH
  // ------------------------------------------

  function escapeHTML(value) {
    return String(value ?? "").replace(/[&<>"']/g, char => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    })[char]);
  }

  function initials(name) {
    return name
      .trim()
      .split(/\s+/)
      .slice(-2)
      .map(word => word[0])
      .join("")
      .toUpperCase();
  }

  function getStatusClass(status) {
    if (status === "Có mặt") return "present";
    if (status === "Đi muộn") return "late";
    return "absent";
  }

  function showToast(title, message) {
    const toast = $("#toast");

    if (!toast) {
      alert(title + "\n" + message);
      return;
    }

    $("#toastTitle").textContent = title;
    $("#toastMessage").textContent = message;
    toast.classList.remove("hidden");

    clearTimeout(showToast.timer);

    showToast.timer = setTimeout(() => {
      toast.classList.add("hidden");
    }, 3000);
  }

  // ------------------------------------------
  // 3. HIỂN THỊ NGÀY
  // ------------------------------------------

  function setDateLabels() {
    const now = new Date();

    const fullDate = $("#fullDate");
    const todayLabel = $("#todayLabel");

    if (fullDate) {
      fullDate.textContent = new Intl.DateTimeFormat(
        "vi-VN",
        {
          weekday: "long",
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }
      ).format(now);
    }

    if (todayLabel) {
      todayLabel.textContent = new Intl.DateTimeFormat(
        "vi-VN",
        {
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }
      ).format(now);
    }
  }

  // ------------------------------------------
  // 4. THỐNG KÊ
  // ------------------------------------------

  function updateStats() {
    const present = rows.filter(
      row => row.status === "Có mặt"
    ).length;

    const late = rows.filter(
      row => row.status === "Đi muộn"
    ).length;

    const absent = rows.filter(
      row => row.status === "Vắng mặt"
    ).length;

    const values = {
      "#totalEmployees": Math.max(25, rows.length),
      "#presentCount": present,
      "#lateCount": late,
      "#absentCount": absent,
      "#legendPresent": present,
      "#legendLate": late,
      "#legendAbsent": absent
    };

    Object.entries(values).forEach(([selector, value]) => {
      const element = $(selector);
      if (element) element.textContent = value;
    });

    const rate = rows.length
      ? Math.round((present / rows.length) * 100)
      : 0;

    if ($("#attendanceRate")) {
      $("#attendanceRate").textContent = rate + "%";
    }

    const donut = $(".donut");

    if (donut) {
      const green = rows.length
        ? (present / rows.length) * 100
        : 0;

      const orange = rows.length
        ? (late / rows.length) * 100
        : 0;

      donut.style.background =
        `conic-gradient(
          #20b483 0 ${green}%,
          #f1b64e ${green}% ${green + orange}%,
          #ed6971 ${green + orange}% 100%
        )`;
    }
  }

  // ------------------------------------------
  // 5. TÌM KIẾM VÀ HIỂN THỊ BẢNG
  // ------------------------------------------

  function getFilteredRows() {
    return rows.filter(row => {
      const matchesFilter =
        activeFilter === "all" ||
        row.status === activeFilter;

      const text = [
        row.name,
        row.code,
        row.department
      ].join(" ").toLocaleLowerCase("vi");

      const matchesSearch =
        text.includes(searchTerm.toLocaleLowerCase("vi"));

      return matchesFilter && matchesSearch;
    });
  }

  function renderRows() {
    const tbody = $("#attendanceRows");

    if (!tbody) return;

    const filtered = getFilteredRows();

    tbody.innerHTML = filtered.map(row => `
      <tr>
        <td>
          <div class="employee-cell">
            <div class="employee-avatar">
              ${escapeHTML(initials(row.name))}
            </div>
            <div>
              <div class="employee-name">
                ${escapeHTML(row.name)}
              </div>
              <div class="employee-code">
                ${escapeHTML(row.code)}
              </div>
            </div>
          </div>
        </td>

        <td>
          <span class="department">
            ${escapeHTML(row.department || "—")}
          </span>
        </td>

        <td>${escapeHTML(row.checkIn || "—")}</td>

        <td>${escapeHTML(row.checkOut || "—")}</td>

        <td>
          <span class="status ${getStatusClass(row.status)}">
            ${escapeHTML(row.status)}
          </span>
        </td>

        <td>
          <div class="row-actions">
            <button
              type="button"
              class="row-action"
              data-action="edit"
              data-id="${row.id}"
              title="Sửa"
            >✎</button>

            <button
              type="button"
              class="row-action delete"
              data-action="delete"
              data-id="${row.id}"
              title="Xóa"
            >×</button>
          </div>
        </td>
      </tr>
    `).join("");

    const emptyState = $("#emptyState");

    if (emptyState) {
      emptyState.classList.toggle(
        "hidden",
        filtered.length > 0
      );
    }

    if ($("#resultCount")) {
      $("#resultCount").textContent =
        `Hiển thị ${filtered.length} bản ghi mẫu`;
    }

    if ($("#allTabCount")) {
      $("#allTabCount").textContent = rows.length;
    }

    updateStats();
  }

  // ------------------------------------------
  // 6. THÊM VÀ SỬA CHẤM CÔNG
  // ------------------------------------------

  function openModal(row = null) {
    const form = $("#attendanceForm");

    if (!form) return;

    editingId = row ? row.id : null;
    form.reset();

    $("#modalTitle").textContent = row
      ? "Sửa thông tin chấm công"
      : "Thêm chấm công";

    if (row) {
      form.elements.name.value = row.name;
      form.elements.code.value = row.code;
      form.elements.department.value = row.department;
      form.elements.checkIn.value =
        row.checkIn === "—" ? "" : row.checkIn;
      form.elements.checkOut.value =
        row.checkOut === "—" ? "" : row.checkOut;
      form.elements.status.value = row.status;
    }

    $("#modalOverlay").classList.remove("hidden");

    setTimeout(() => {
      form.elements.name.focus();
    }, 50);
  }

  function closeModal() {
    const overlay = $("#modalOverlay");

    if (overlay) {
      overlay.classList.add("hidden");
    }

    editingId = null;
  }

  const attendanceForm = $("#attendanceForm");

  if (attendanceForm) {
    attendanceForm.addEventListener("submit", event => {
      event.preventDefault();

      const form = event.currentTarget;

      const name = form.elements.name.value.trim();
      const code = form.elements.code.value.trim().toUpperCase();
      const department = form.elements.department.value.trim();

      const duplicate = rows.some(row =>
        row.code.toLowerCase() === code.toLowerCase() &&
        row.id !== editingId
      );

      if (duplicate) {
        showToast(
          "Mã nhân viên đã tồn tại",
          "Vui lòng nhập một mã nhân viên khác."
        );
        return;
      }

      const data = {
        id: editingId ?? (
          Math.max(0, ...rows.map(row => row.id)) + 1
        ),
        name,
        code,
        department,
        checkIn: form.elements.checkIn.value || "—",
        checkOut: form.elements.checkOut.value || "—",
        status: form.elements.status.value
      };

      if (editingId !== null) {
        rows = rows.map(row =>
          row.id === editingId ? data : row
        );
      } else {
        rows.unshift(data);
      }

      const wasEditing = editingId !== null;

      closeModal();
      renderRows();

      showToast(
        wasEditing ? "Đã sửa bản ghi mẫu" : "Đã thêm bản ghi mẫu",
        "Thay đổi chưa được lưu lên máy chủ."
      );
    });
  }

  if ($("#addAttendanceBtn")) {
    $("#addAttendanceBtn").addEventListener(
      "click",
      () => openModal()
    );
  }

  if ($("#closeModal")) {
    $("#closeModal").addEventListener("click", closeModal);
  }

  if ($("#cancelModal")) {
    $("#cancelModal").addEventListener("click", closeModal);
  }

  if ($("#modalOverlay")) {
    $("#modalOverlay").addEventListener("click", event => {
      if (event.target === $("#modalOverlay")) {
        closeModal();
      }
    });
  }

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeModal();
  });

  // ------------------------------------------
  // 7. XÓA BẢN GHI
  // ------------------------------------------

  const attendanceRows = $("#attendanceRows");

  if (attendanceRows) {
    attendanceRows.addEventListener("click", event => {
      const button = event.target.closest("[data-action]");

      if (!button) return;

      const id = Number(button.dataset.id);
      const row = rows.find(item => item.id === id);

      if (!row) return;

      if (button.dataset.action === "edit") {
        openModal(row);
      }

      if (button.dataset.action === "delete") {
        const confirmed = confirm(
          `Bạn có chắc muốn xóa bản ghi của ${row.name}?`
        );

        if (!confirmed) return;

        rows = rows.filter(item => item.id !== id);

        renderRows();

        showToast(
          "Đã xóa bản ghi mẫu",
          "Dữ liệu chưa được lưu lên máy chủ."
        );
      }
    });
  }

  // ------------------------------------------
  // 8. LỌC VÀ TÌM KIẾM
  // ------------------------------------------

  $$(".filter-tab").forEach(button => {
    button.addEventListener("click", () => {
      activeFilter = button.dataset.filter;

      $$(".filter-tab").forEach(tab => {
        tab.classList.toggle("selected", tab === button);
      });

      renderRows();
    });
  });

  if ($("#globalSearch")) {
    $("#globalSearch").addEventListener("input", event => {
      searchTerm = event.target.value.trim();
      renderRows();
    });
  }

  // ------------------------------------------
  // 9. CHUYỂN MENU
  // ------------------------------------------

  const pages = {
    dashboard: [
      "Tổng quan",
      "Chào bạn, Admin",
      "Theo dõi tình hình chấm công và nhân sự của bạn."
    ],
    attendance: [
      "Chấm công",
      "Quản lý chấm công",
      "Theo dõi giờ vào, giờ ra và trạng thái nhân viên."
    ],
    employees: [
      "Nhân viên",
      "Quản lý nhân viên",
      "Danh sách nhân sự minh họa."
    ],
    upload: [
      "Tải file lên",
      "Tải bảng số liệu lên",
      "Chọn file Excel hoặc CSV để xem trước."
    ],
    reports: [
      "Thống kê & báo cáo",
      "Thống kê & báo cáo",
      "Các chỉ số đang hiển thị là dữ liệu mẫu."
    ],
    settings: [
      "Cài đặt",
      "Cài đặt hệ thống",
      "Các tùy chỉnh sẽ được bổ sung sau."
    ]
  };

  function navigate(page) {
    const details = pages[page] || pages.dashboard;

    $("#breadcrumbTitle").textContent = details[0];
    $("#pageTitle").textContent = details[1];
    $("#pageSubtitle").textContent = details[2];

    $$(".nav-link").forEach(button => {
      button.classList.toggle(
        "active",
        button.dataset.page === page
      );
    });

    $("#sidebar").classList.remove("open");

    const isUpload = page === "upload";

    $(".app-shell").classList.toggle(
      "page-upload",
      isUpload
    );

    if ($("#uploadPage")) {
      $("#uploadPage").classList.toggle(
        "hidden",
        !isUpload
      );
    }

    if (isUpload) {
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    if (page === "attendance" || page === "employees") {
      $("#attendancePanel").scrollIntoView({
        behavior: "smooth",
        block: "start"
      });
    } else if (page === "reports") {
      $(".summary-panel").scrollIntoView({
        behavior: "smooth",
        block: "center"
      });
    } else if (page === "settings") {
      showToast(
        "Cài đặt",
        "Mục này sẽ được phát triển sau."
      );
    } else {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }

  $$("[data-page]").forEach(button => {
    button.addEventListener("click", () => {
      navigate(button.dataset.page);
    });
  });

  if ($("#menuToggle")) {
    $("#menuToggle").addEventListener("click", () => {
      $("#sidebar").classList.toggle("open");
    });
  }

  // ------------------------------------------
  // 10. XUẤT CSV
  // ------------------------------------------

  if ($("#exportBtn")) {
    $("#exportBtn").addEventListener("click", () => {
      const filtered = getFilteredRows();

      const csvCell = value =>
        '"' + String(value ?? "").replace(/"/g, '""') + '"';

      const csv = [
        [
          "Mã nhân viên",
          "Tên nhân viên",
          "Phòng ban",
          "Giờ vào",
          "Giờ ra",
          "Trạng thái"
        ].map(csvCell).join(","),

        ...filtered.map(row => [
          row.code,
          row.name,
          row.department,
          row.checkIn,
          row.checkOut,
          row.status
        ].map(csvCell).join(","))
      ].join("\r\n");

      const blob = new Blob(
        ["\uFEFF", csv],
        { type: "text/csv;charset=utf-8;" }
      );

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");

      link.href = url;
      link.download = "du-lieu-cham-cong-mau.csv";

      document.body.appendChild(link);
      link.click();
      link.remove();

      setTimeout(() => URL.revokeObjectURL(url), 1000);

      showToast(
        "Đã xuất CSV",
        "Tệp dữ liệu mẫu đã được tải xuống."
      );
    });
  }

  // ------------------------------------------
  // 11. ĐỌC FILE EXCEL VÀ CSV
  // ------------------------------------------

  const fileInput = $("#fileInput");
  const uploadMessage = $("#uploadMessage");
  const uploadPreview = $("#uploadPreview");

  function escapeCell(value) {
    return escapeHTML(value);
  }

  function showFilePreview(matrix, fileName) {
    if (!matrix.length) {
      uploadMessage.textContent = "File không có dữ liệu.";
      return;
    }

    const headers = matrix[0];

    $("#uploadFileName").textContent = fileName;

    $("#uploadHead").innerHTML =
      "<tr>" +
      headers.map((value, index) =>
        `<th>${escapeCell(value || "Cột " + (index + 1))}</th>`
      ).join("") +
      "</tr>";

    $("#uploadBody").innerHTML =
      matrix.slice(1, 21).map(row =>
        "<tr>" +
        headers.map((_, index) =>
          `<td>${escapeCell(row[index] ?? "")}</td>`
        ).join("") +
        "</tr>"
      ).join("");

    uploadPreview.classList.remove("hidden");

    uploadMessage.textContent =
      `Đọc thành công ${Math.max(matrix.length - 1, 0)} dòng. ` +
      "Đang hiển thị tối đa 20 dòng đầu tiên.";
  }

  if (fileInput) {
    fileInput.addEventListener("change", async function () {
      const file = this.files[0];

      if (!file) return;

      const extension = file.name
        .split(".")
        .pop()
        .toLowerCase();

      if (!["xlsx", "xls", "csv"].includes(extension)) {
        uploadMessage.textContent =
          "Vui lòng chọn file Excel hoặc CSV.";
        this.value = "";
        return;
      }

      if (file.size > 15 * 1024 * 1024) {
        uploadMessage.textContent =
          "File vượt quá giới hạn 15 MB.";
        this.value = "";
        return;
      }

      try {
        if (!window.XLSX) {
          throw new Error(
            "Không tải được thư viện Excel. Hãy kiểm tra Internet."
          );
        }

        const buffer = await file.arrayBuffer();

        const workbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true
        });

        const sheet =
          workbook.Sheets[workbook.SheetNames[0]];

        const matrix = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
          raw: false
        });

        showFilePreview(matrix, file.name);

      } catch (error) {
        uploadMessage.textContent =
          error.message ||
          "Không đọc được file. Hãy thử file khác.";
      }

      this.value = "";
    });
  }

  // ------------------------------------------
  // 12. KHỞI ĐỘNG WEBSITE
  // ------------------------------------------

  setDateLabels();
  renderRows();

})();
