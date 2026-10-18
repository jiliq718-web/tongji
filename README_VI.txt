HƯỚNG DẪN CÀI ĐẶT WEBSITE THỐNG KÊ ĐẠI LÝ

A. Cấu hình Supabase
1. Mở đúng dự án Supabase bạn muốn dùng.
2. Vào SQL Editor > New query.
3. Mở file supabase.sql, sao chép toàn bộ nội dung và chạy (Run).
4. Vào Project Settings > API (hoặc API Keys).
5. Sao chép Project URL và anon/public key.
6. Mở app.js:
   - Thay PASTE_SUPABASE_PROJECT_URL_HERE bằng Project URL.
   - Thay PASTE_SUPABASE_ANON_KEY_HERE bằng anon/public key.
   - Không bao giờ đưa service_role/secret key vào website.

B. Tạo tài khoản
1. Đẩy code lên GitHub rồi chờ Netlify deploy.
2. Mở website, nhập email và mật khẩu rồi bấm “Tạo tài khoản”.
3. Nếu Supabase yêu cầu xác nhận email, hãy mở email và xác nhận.
4. Đăng nhập bằng tài khoản đó.
5. Nếu bạn chỉ muốn một số người dùng được xem dữ liệu, hãy chỉ tạo tài khoản cho những người đó.

C. Đưa code lên GitHub
1. Mở repository của website.
2. Thay nội dung index.html, style.css, app.js bằng các file cùng tên trong gói này.
3. Thêm file supabase.sql và README_VI.txt (có thể không đưa README lên website cũng được).
4. Commit changes. Netlify sẽ tự deploy nếu đã liên kết repository.
5. Mở website Netlify và đăng nhập.

D. Nhập Excel
1. File Excel cần có hàng tiêu đề. Công cụ dò tiêu đề trong 20 dòng đầu.
2. Tên cột nên tương ứng với ảnh Excel gốc (tiếng Trung) hoặc tiêu đề tiếng Việt trong index.html/app.js.
3. Dữ liệu ngày dạng ngày Excel, YYYY-MM-DD, hoặc dạng “10月1号”.
4. Các dòng “总计” được bỏ qua; website tự tính tổng từ các dòng dữ liệu đang lọc.
5. Nếu file có ô mã đại lý gộp nhiều dòng, công cụ sẽ tự điền tiếp mã đại lý xuống các dòng trống.
6. Hiện tại khóa cập nhật là mã đại lý + ngày. Nếu cùng một đại lý có nhiều dòng hợp lệ trong cùng ngày, cần điều chỉnh khóa dữ liệu trước khi nhập.

E. Lưu ý bảo mật và dữ liệu
- Website dùng Supabase Auth; bảng bật Row Level Security, người chưa đăng nhập không đọc/ghi được.
- Người dùng đã đăng nhập có cùng quyền trên bảng. Nếu cần phân quyền theo từng tài khoản, phải bổ sung bảng phân quyền và policy.
- Sao lưu file Excel gốc trước khi nhập.
- Một số tiêu đề trong ảnh Trung Quốc hơi khó phân biệt do độ phân giải. Nếu cột nào bị lệch, gửi file Excel gốc để điều chỉnh ánh xạ chính xác.
