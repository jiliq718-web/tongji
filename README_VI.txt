HUONG DAN THEM CHUC NANG CAP TREN / CAP DUOI

1) Sao luu code hien tai tren GitHub truoc khi thay file.
2) Trong Supabase > SQL Editor > New query, mo file supabase_hierarchy.sql, sao chep toan bo va Run.
   File nay chi tao bang agent_hierarchy; khong xoa hay sua du lieu agent_stats.
3) Tren GitHub repository jiliq718-web/tongji, thay noi dung index.html, style.css, app.js bang cac file cung ten trong goi nay, sau do Commit changes.
   app.js da giu lai Project URL va Publishable key da xac nhan trong cuoc tro chuyen.
4) Cho Netlify deploy xong, tai lai website (Ctrl+F5) va dang nhap.
5) Mo tab Quan he cap tren. Nhap ma dai ly cap duoi va ma cap tren truc tiep, bam Luu quan he.
6) Lap lai cho tung quan he. Moi dai ly chi co mot cap tren truc tiep.
7) Sang tab Thong ke du lieu, nhap dung ma cap tren vao o tim kiem. Website se hien ma do va tat ca cap duoi qua nhieu tang.

Vi du:
- 10001 co cap tren 10002
- 10002 co cap tren 10003
Tim 10003 se hien 10003, 10002, 10001 (neu cac ma co du lieu thong ke).

Luu y:
- Neu chua co bang agent_hierarchy, chuc nang quan he cap tren se bao loi cho den khi chay SQL migration.
- RLS chi cho phep tai khoan da dang nhap. Trong code hien tai, moi tai khoan da dang nhap co quyen xem/sua chung du lieu.
- Khong dua secret/service_role key vao app.js.
