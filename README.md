# Chấm Công Workspace — Giao diện v1

Website giao diện quản lý chấm công bằng HTML, CSS và JavaScript thuần. Không cần cài Node.js hoặc chạy lệnh build.

## Chạy thử trên Windows
1. Giải nén tệp ZIP.
2. Mở thư mục `quan-ly-cham-cong-v1`.
3. Nhấp đúp vào `index.html` để mở bằng Chrome hoặc Edge.

## Đưa lên GitHub
1. Tạo repository mới trên GitHub.
2. Tải lên `index.html`, `style.css`, `app.js` và `README.md` (có thể kéo thả các tệp vào trang repository).
3. Commit changes.

## Triển khai trên Netlify
1. Đăng nhập Netlify.
2. Chọn Add new project → Import an existing project.
3. Kết nối GitHub và chọn repository.
4. Để trống Build command; Publish directory là `.` nếu các tệp nằm ngay ở thư mục gốc.
5. Deploy.

## Tính năng hiện có
- Giao diện responsive: desktop, tablet và điện thoại.
- Sidebar điều hướng.
- Bảng dữ liệu demo, tìm kiếm và lọc trạng thái.
- Thêm, sửa, xóa bản ghi trong bộ nhớ của phiên làm việc.
- Xuất CSV.
- Thống kê minh họa.

## Chưa có trong v1
- Đăng nhập thật và phân quyền.
- Cơ sở dữ liệu dùng chung.
- Đồng bộ giữa nhiều người hoặc nhiều thiết bị.
- Nhập Excel `.xlsx`.
- Lọc chấm công theo ngày thật.

Các bản ghi mẫu và thay đổi thêm/sửa/xóa chỉ nằm trong bộ nhớ JavaScript, không được lưu lâu dài. Khi thêm tính năng dùng chung, có thể kết nối Supabase và bật Row Level Security (RLS).
