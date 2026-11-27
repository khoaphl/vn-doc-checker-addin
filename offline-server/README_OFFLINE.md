# HƯỚNG DẪN CÀI ĐẶT & CHẠY CHẾ ĐỘ 100% OFFLINE (PC ĐỒNG THÁP)

Chế độ này giúp máy tính **mất mạng hoàn toàn ngay từ lúc bật máy** vẫn mở và sử dụng được Word Add-in **Rà soát văn bản** (Kiểm tra Bước 1, Chuẩn hóa NĐ 30, và Kiểm tra chính tả Offline 3 lớp).

## Bước 1: Yêu cầu môi trường
- Đã cài đặt **Node.js** (v18 trở lên). Kiểm tra bằng cách mở Command Prompt / PowerShell gõ `node -v`.

## Bước 2: Tạo & tin cậy chứng chỉ SSL `localhost` (làm 1 lần duy nhất)
1. Mở PowerShell tại thư mục dự án `vn-doc-checker-addin`.
2. Chạy lệnh:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\offline-server\certs\create-cert.ps1
   ```
3. Khi Windows hiện hộp thoại **Security Warning** hỏi có cài chứng chỉ vào *Trusted Root Certification Authorities* không, bấm **Yes**.

## Bước 3: Cài đặt chạy ngầm tự động khi bật Windows
1. Nhấn `Win + R`, gõ `shell:startup` rồi nhấn **Enter**.
2. Chuột phải vào file `offline-server\start-hidden.vbs` → chọn **Create shortcut** (Tạo lối tắt).
3. Di chuyển shortcut vừa tạo vào thư mục **Startup** vừa mở ở bước 1.
4. Nhấp đúp `start-hidden.vbs` để khởi chạy ngay lập tức (chạy ngầm hoàn toàn, không hiện cửa sổ đen).
5. Kiểm tra nhanh bằng trình duyệt: mở `https://localhost:8443/taskpane.html` — nếu thấy giao diện Add-in mà không báo lỗi chứng chỉ là thành công.

## Bước 4: Nạp `manifest-offline.xml` vào Microsoft Word
1. Tạo một thư mục dùng chung trên máy (ví dụ `C:\WordAddins`) và chép file `manifest-offline.xml` vào đó, sau đó bật **Share** thư mục đó cho chính tài khoản Windows của bạn (Read).
2. Mở Microsoft Word → **File** → **Options** → **Trust Center** → **Trust Center Settings...** → **Trusted Add-in Catalogs**.
3. Dán đường dẫn mạng của thư mục vừa share (VD: `\\TEN-MAY-TINH\WordAddins`) vào ô **Catalog Url** → bấm **Add catalog** → tích chọn **Show in Menu** → **OK** và khởi động lại Word.
4. Trong Word, vào **Insert** (hoặc **Home**) → **Add-ins** → **Shared Folder** → chọn **Rà soát văn bản Offline (PC Đồng Tháp)**.

> **Lưu ý:** File `manifest.xml` ở thư mục gốc vẫn giữ nguyên cấu hình Online để đẩy lên GitHub Pages (`https://khoaphl.github.io/vn-doc-checker-addin/`).
