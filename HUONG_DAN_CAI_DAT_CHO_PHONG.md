# HƯỚNG DẪN TRIỂN KHAI WORD ADD-IN "RÀ SOÁT VĂN BẢN (PC ĐỒNG THÁP)" CHO THÀNH VIÊN TRONG PHÒNG

Tài liệu này hướng dẫn 2 cách triển khai Add-in cho các máy tính trong phòng:
- **Cách 1 (Khuyên dùng — Online qua GitHub Pages):** Dễ cài nhất, khi có bản cập nhật mới thì **tất cả máy tự động cập nhật**, không cần cài lại.
- **Cách 2 (100% Offline):** Dành cho máy tính không có Internet hoặc bị chặn mạng ngoài.

---

## CÁCH 1: CÀI ĐẶT BẢN ONLINE (`manifest.xml`) — KHUYÊN DÙNG

### Bước 1: Chuẩn bị nguồn phát hành (Người quản lý làm 1 lần)
1. Đẩy toàn bộ mã nguồn dự án lên GitHub và bật **GitHub Pages** (hiện tại đã cấu hình sẵn tại `https://khoaphl.github.io/vn-doc-checker-addin/`).
2. Chọn 1 trong 2 cách phân phối file `manifest.xml` cho các thành viên:
   - **Phương án A (Dùng thư mục chia sẻ LAN — Nhanh nhất):**
     1. Trên máy chủ nội bộ (hoặc máy trưởng phòng/chuyên viên phụ trách), tạo thư mục `C:\WordAddins` và chép file `manifest.xml` vào đó.
     2. Chuột phải vào thư mục `WordAddins` → **Properties** → tab **Sharing** → **Share...** → thêm **Everyone** (hoặc nhóm nhân viên trong phòng) với quyền **Read** → bấm **Share**.
     3. Ghi lại đường dẫn mạng, ví dụ: `\\MAY-CHU-PHONG\WordAddins`.
   - **Phương án B (Mỗi máy tự lưu cục bộ — Nếu mạng LAN không share được):**
     - Gửi file `manifest.xml` qua Zalo/Email/D-Office cho từng thành viên.
     - Trên máy thành viên, tạo `C:\WordAddins`, chép `manifest.xml` vào đó và bật **Share** cho chính tài khoản của máy đó → lấy đường dẫn `\\TEN-MAY-CA-NHAN\WordAddins`.

> **Mẹo (Nếu cơ quan dùng Microsoft 365 quản trị tập trung):** Có thể gửi file `manifest.xml` cho bộ phận IT (Phòng VTCNTT) để đẩy qua **Microsoft 365 Admin Center → Integrated apps**. Khi đó Word của tất cả mọi người sẽ tự hiện nút Add-in mà không cần thao tác tay bên dưới.

---

### Bước 2: Gắn Add-in vào Microsoft Word (Làm 1 lần trên từng máy thành viên)
1. Mở **Microsoft Word** (mở 1 tài liệu trống bất kỳ).
2. Vào **File** → **Options** → **Trust Center** → bấm **Trust Center Settings...**
3. Chọn mục **Trusted Add-in Catalogs** ở cột bên trái.
4. Tại ô **Catalog Url**, nhập đường dẫn thư mục chia sẻ ở Bước 1 (ví dụ `\\MAY-CHU-PHONG\WordAddins`) → bấm **Add catalog**.
5. Ở bảng danh sách phía dưới, **tích chọn ô `Show in Menu`** tại dòng vừa thêm → bấm **OK** → **OK**.
6. **Tắt hoàn toàn Microsoft Word** (đóng tất cả cửa sổ Word) rồi mở lại.
7. Trên thanh công cụ Word, vào thẻ **Home** (hoặc **Insert**) → bấm **Add-ins** (hoặc **Get Add-ins / My Add-ins**) → chọn tab **Shared Folder** → chọn **Rà soát văn bản (PC Đồng Tháp)** → bấm **Add**.
8. Từ lúc này, trên thẻ **Home** góc phải sẽ xuất hiện nhóm **PC Đồng Tháp** với nút **Rà soát văn bản**.

---

## CÁCH 2: CÀI ĐẶT BẢN 100% OFFLINE (`manifest-offline.xml`)

Áp dụng cho máy mất mạng hoàn toàn hoặc không truy cập được GitHub Pages:

1. Chép toàn bộ thư mục `vn-doc-checker-addin` vào ổ đĩa trên máy thành viên (ví dụ `D:\vn-doc-checker-addin`).
2. Cài đặt **Node.js** (bản LTS) nếu máy chưa có.
3. Mở PowerShell tại thư mục `D:\vn-doc-checker-addin`, chạy lệnh tạo chứng chỉ SSL cục bộ (1 lần duy nhất):
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\offline-server\certs\create-cert.ps1
   ```
   *(Khi Windows hỏi xác nhận tin cậy chứng chỉ, bấm **Yes**)*.
4. Cho phép server tự chạy ngầm khi bật Windows:
   - Nhấn `Win + R` → gõ `shell:startup` → **Enter**.
   - Tạo Shortcut của file `D:\vn-doc-checker-addin\offline-server\start-hidden.vbs` và chép vào thư mục Startup vừa mở.
   - Nhấp đúp `start-hidden.vbs` để khởi chạy ngay.
5. Nạp file `manifest-offline.xml` vào Word theo đúng các bước ở **Bước 2 (Cách 1)** phía trên (lưu ý chọn **Rà soát văn bản Offline (PC Đồng Tháp)**).

---

## CẤU HÌNH SAU KHI CÀI ĐẶT TRÊN MÁY THÀNH VIÊN

1. **Google Gemini API Key (để AI rà soát Bước 2 Online & Bước 3):**
   - Thành viên tự tạo key miễn phí tại [Google AI Studio](https://aistudio.google.com/) hoặc dùng key nội bộ do phòng cấp, dán vào ô **Google Gemini API Key (BYOK)** rồi bấm **Lưu API Key**.
   - *Lưu ý:* Nếu không nhập API Key hoặc bị mất mạng, Add-in vẫn chạy bình thường **Bước 1 (Kỹ thuật & D-Office)**, **Chuẩn hóa NĐ 30**, và **Bước 2 (Kiểm tra chính tả Offline 3 lớp)**.
2. **Đồng bộ Quy định nội bộ & Danh sách bỏ qua chung của phòng:**
   - Gửi kèm file `rules_pcdt.txt` để thành viên nạp vào mục **Quy định riêng nội bộ**.
   - Nếu phòng có danh sách các từ viết tắt/tên riêng ngoại lệ hay dùng, 1 người bấm **📥 Xuất file `.txt`** trong mục **Quản lý danh sách bỏ qua chung**, gửi file cho cả phòng bấm **📤 Nhập file `.txt`** (hệ thống tự gộp thêm và lọc trùng).
3. **Lưu thông số NĐ 30 mặc định:**
   - Mở **📐 Kiểm tra & Chuẩn hóa (NĐ 30)**, chọn sẵn 8 thông số theo thống nhất của phòng (VD: cỡ chữ `13 pt`, thụt đầu dòng `1,27 cm`, giãn dòng `1,15`, cách đoạn Before & After `6 pt`, lề chuẩn) rồi bấm **💾 Lưu làm mặc định**.
