# ĐẶC TẢ YÊU CẦU SẢN PHẨM & KỸ THUẬT (PRD / TECHNICAL SPEC)
**Dự án:** Word Add-in Rà soát văn bản hành chính — Công ty Điện lực Đồng Tháp (PCĐT)  
**Tên kho lưu trữ:** `vn-doc-checker-addin`  
**Phiên bản mục tiêu:** v13 (Hỗ trợ Online + 100% Offline, Chuẩn hóa NĐ 30 đồng bộ Before/After)

---

## 1. TỔNG QUAN & KIẾN TRÚC HỆ THỐNG

### 1.1. Mục tiêu
Xây dựng Microsoft Word Add-in hỗ trợ cán bộ, chuyên viên tại Công ty Điện lực Đồng Tháp tự động kiểm tra kỹ thuật trình bày, chính tả, thể thức và tính hợp lệ pháp lý của văn bản hành chính theo:
- Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư.
- Quyết định số 739/QĐ-EVN ngày 10/5/2022 và Quyết định số 1080/QĐ-EVN ngày 01/8/2021 của Tập đoàn Điện lực Việt Nam (EVN).
- Quyết định số 08/QĐ-HĐTV ngày 12/01/2026 của HĐTV Tổng công ty Điện lực miền Nam (EVNSPC).
- Công văn số 3156/PCĐT-VP ngày 20/4/2026 và Công văn số 9975/PCĐT-VP ngày 03/11/2026 của Công ty Điện lực Đồng Tháp.

### 1.2. Công nghệ & Chế độ hoạt động kép (Online + 100% Offline)
- **Ngôn ngữ:** HTML5, CSS3, JavaScript (ES6+) kết hợp Microsoft `Office.js` (không chuyển sang Rust/WebAssembly).
- **Chế độ Online (GitHub Pages):** Triển khai tại `https://khoaphl.github.io/vn-doc-checker-addin/` thông qua `manifest.xml`.
- **Chế độ 100% Offline (mất mạng ngay từ lúc bật máy):**
  - Đóng gói kèm giải pháp chạy máy chủ HTTPS cục bộ (`localhost`) chạy ngầm trên Windows.
  - Lưu trữ cục bộ toàn bộ tài nguyên tĩnh (HTML, CSS, JS, icon, bộ từ điển JSON) và bản sao cục bộ của thư viện `office.js` để Word có thể nạp giao diện Add-in ngay cả khi máy tính không có kết nối internet.
  - Cung cấp file `manifest-offline.xml` riêng trỏ về địa chỉ `https://localhost:<port>/taskpane.html`.

### 1.3. Tích hợp giao diện Microsoft Word (`manifest.xml` & `manifest-offline.xml`)
- Khai báo `VersionOverrides` gắn trực tiếp vào thẻ **Home** (`TabHome`).
- Tên nhóm (Group): **PC Đồng Tháp**.
- Tên nút (Button): **Rà soát văn bản**.
- Bộ biểu tượng (`assets/icon-16.png`, `icon-32.png`, `icon-80.png`): Kết hợp biểu tượng rà soát văn bản cùng ngôi sao 4 cánh EVN đúng mã màu nhận diện thương hiệu (`RGB 22,67,151`, `RGB 237,28,36`, `RGB 255,242,0`).

---

## 2. CẤU TRÚC THƯ MỤC ĐỀ XUẤT

Tách mã nguồn thành các module rõ ràng để dễ kiểm soát lỗi cú pháp và bảo trì:

```text
vn-doc-checker-addin/
├── manifest.xml                  # Bản Online (trỏ về GitHub Pages)
├── manifest-offline.xml          # Bản Offline (trỏ về https://localhost)
├── taskpane.html                 # Giao diện chính của Add-in
├── rules_pcdt.txt                # Bộ quy định nội bộ hợp nhất của PC Đồng Tháp
├── js/
│   ├── app.js                    # Khởi tạo Office.onReady, điều phối UI & sự kiện
│   ├── word-utils.js             # Đọc văn bản sạch, tìm kiếm trong Table, Comment, Track Changes
│   ├── nd30-formatter.js         # Quét & Chuẩn hóa 7 tiêu chí NĐ 30 (có xác nhận & lưu mặc định)
│   ├── checker-step1.js          # Bước 1: Kiểm tra kỹ thuật & thể thức cứng (Office.js + Regex)
│   ├── checker-ai.js             # Gọi Gemini API (Bước 2 & Bước 3, retry & fallback model)
│   ├── spellcheck-offline.js     # Bộ kiểm tra chính tả ngoại tuyến 3 lớp
│   └── storage-manager.js        # Quản lý API Key, rules_pcdt.txt, Global Ignore, Từ điển riêng
├── data/
│   ├── vn-syllables.json         # Bộ ~7.000 âm tiết tiếng Việt hợp lệ
│   ├── common-typos.json         # Danh sách cặp từ hay nhầm, khẩu ngữ/biệt ngữ
│   └── evn-dict.json             # Từ viết tắt ngành Điện & địa danh tỉnh Đồng Tháp
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── offline-server/
    ├── server.js (hoặc script)   # Máy chủ HTTPS cục bộ nhẹ cho Windows
    ├── certs/                    # Hướng dẫn/script tự tạo chứng chỉ localhost tin cậy
    └── office-js-local/          # Thư viện office.js lưu cục bộ dùng khi mất mạng
```

---

## 3. ĐẶC TẢ CHỨC NĂNG CHI TIẾT

### 3.1. Tiền xử lý văn bản & Nhận diện loại văn bản
1. **Đọc văn bản sạch (`extractCleanDocumentText`):**
   - Duyệt từng đoạn (`context.document.body.paragraphs`), chuyển mọi ký tự ngắt dòng mềm (`Shift + Enter`, `\v`, `\u000b`, `\r`) thành dấu xuống dòng thực sự (`\n`).
   - Mục đích: Khắc phục triệt để lỗi dính chữ khi văn bản nằm trong ô bảng (Table) hoặc xuống dòng trong tiêu đề (ví dụ tránh lỗi đọc nhầm thành `TỔNG CÔNG TYĐIỆN LỰC MIỀN NAM`).
2. **Tự động nhận diện loại văn bản:**
   - Quét 20 dòng đầu bằng Regex để nhận diện nhanh: `Quyết định`, `Tờ trình`, `Báo cáo`, `Thông báo`, `Kế hoạch`, `Biên bản`, `Quy định / Quy chế`, hoặc dòng bắt đầu bằng `V/v` → `Công văn`.
   - Nếu chưa xác định được và đang có mạng, gọi Gemini AI đọc 1.200 ký tự đầu để phân loại.
   - Cho phép người dùng chọn thủ công lại loại văn bản qua menu thả xuống.

### 3.2. Quy trình kiểm tra 3 bước tuần tự

#### Bước 1: Kiểm tra Kỹ thuật & Thể thức cứng (100% cục bộ, không gọi AI)
- **Căn lề trang (NĐ 30):** Trên/Dưới `20–25 mm`, Trái `30–35 mm`, Phải `15–20 mm`.
- **Phông chữ & Cỡ chữ:** Bắt lỗi phông khác `Times New Roman`, cỡ chữ nội dung ngoài khoảng `12–14 pt` (cho đoạn dài > 40 ký tự).
- **Quy tắc cứng D-Office EVN & PC Đồng Tháp:**
  - Thiếu cụm từ `Số:` (để D-Office cấp số tự động).
  - Thiếu ký hiệu kết thúc `./.` (để D-Office chèn chữ ký tắt).
  - Viết tắt viện dẫn sai quy định (`CV1041`, `TB134`, `VB294`...).
  - Riêng **Công văn:** Trích yếu lồng tên loại (`V/v thông báo...`, `V/v báo cáo...`).
  - 15 dòng cuối văn bản: Viết tắt hoặc ghi sai chức vụ người ký (`PGĐ`, `KT. PGĐ`, `PHÓ ĐIỆN LỰC`).

#### Bước 2: Kiểm tra Chính tả & Ngôn ngữ (Online AI + Tự động dự phòng Offline)
- **Chế độ Online (Gemini API):**
  - Thiết lập `temperature = 0.0`, hỗ trợ xoay vòng model dự phòng và tự retry khi gặp lỗi HTTP `503` / `429`.
  - Phạm vi duy nhất: Chính tả, lỗi đánh máy, dấu câu, khẩu ngữ/biệt ngữ (*vại, bít, thui, chúa nhật...*).
  - Nguyên tắc loại trừ: Không kiểm tra thể thức/lề/font/pháp lý; không báo lỗi thiếu dấu cách giữa 2 dòng khác nhau; không bắt lỗi tên viết tắt ngành Điện (`EVN`, `EVNSPC`, `PCĐT`, `TBA`, `kV`, `MW`...).
  - Tự động chèn *Danh sách bỏ qua chung* (kèm lý do) vào prompt.
  - Trả về trường `replacement` để hỗ trợ nút **✏️ Sửa tự động (Track Changes)**.
- **Chế độ Offline (Tự động kích hoạt ở Bước 2 khi mất mạng hoặc lỗi API):**
  - Hiển thị rõ trạng thái đang chạy chế độ Offline.
  - Kiểm tra theo **3 lớp**:
    1. **Lớp 1 — Âm tiết không tồn tại:** Tách từ và đối chiếu với `data/vn-syllables.json` (~7.000 âm tiết tiếng Việt chuẩn) để phát hiện lỗi gõ sai (VD: *điệ lực, ngiêng, cty...*).
    2. **Lớp 2 — Cặp từ hay nhầm & Khẩu ngữ:** Đối chiếu `data/common-typos.json` (VD: *sát nhập → sáp nhập*, *chuẩn đoán → chẩn đoán*, *vại, bít, thui, chúa nhật...*) kèm gợi ý thay thế (`replacement`).
    3. **Lớp 3 — Từ điển loại trừ & Tự bổ sung:** Bỏ qua các từ có trong `data/evn-dict.json` (thuật ngữ ngành Điện, địa danh Đồng Tháp), *Danh sách bỏ qua chung*, và **Từ điển riêng của người dùng** (cho phép thêm/xóa trực tiếp trên giao diện, lưu trong `localStorage`).

#### Bước 3: Kiểm tra Thể thức & Tính hợp lệ theo loại văn bản (Gemini AI)
- Ưu tiên áp dụng nội dung file `rules_pcdt.txt` (do người dùng nạp vào `localStorage`) và *Danh sách bỏ qua chung*.
- Tuyệt đối không lặp lại lỗi của Bước 1 và Bước 2.
- **Quy tắc bắt buộc đối với Quyết định:**
  - Phân biệt rõ **Vị trí 1 — Thẩm quyền ban hành** (nằm ngay dưới trích yếu và trên phần `Căn cứ...`, VD: `GIÁM ĐỐC CÔNG TY ĐIỆN LỰC ĐỒNG THÁP` — **TUYỆT ĐỐI KHÔNG** báo lỗi thiếu `KT.`/`TM.`, thiếu họ tên hay thiếu dòng ký) với **Vị trí 2 — Người ký ban hành** (cuối văn bản, sau `./.`, ngang hàng `Nơi nhận:`).
  - Đối chiếu chính xác **4 Quyết định nền tảng của PC Đồng Tháp**:
    1. `Quyết định số 235/QĐ-EVN ngày 14 tháng 4 năm 2010` của Tập đoàn Điện lực Việt Nam về việc đổi tên các Điện lực trực thuộc Tổng công ty Điện lực miền Nam *(chính là Quyết định thành lập/đổi tên Công ty, đặt đầu tiên là ĐÚNG)*.
    2. `Quyết định số 201/QĐ-HĐTV ngày 05 tháng 8 năm 2025` của HĐTV Tổng công ty Điện lực miền Nam về việc ban hành Quy chế tổ chức và hoạt động của Công ty Điện lực Đồng Tháp.
    3. `Quyết định số 1677/QĐ-EVNSPC ngày 01 tháng 6 năm 2025` của Tổng công ty Điện lực miền Nam ban hành Quy chế phân cấp quản lý.
    4. `Quyết định số 1824/QĐ-EVNSPC ngày 17 tháng 6 năm 2025` của Tổng công ty Điện lực miền Nam về việc sửa đổi, bổ sung Quy chế phân cấp.
- **Khi mất mạng:** Không treo ứng dụng; hiển thị thông báo đã hoàn tất Bước 1 & Bước 2 (Offline), Bước 3 tạm bỏ qua do không có kết nối mạng.

### 3.3. Tương tác trên thẻ lỗi & Cơ chế Bỏ qua 2 cấp độ
- **Thao tác trên từng thẻ lỗi:**
  - `📍 Đến vị trí`: Tìm kiếm thông minh cả ngoài bảng lẫn sâu trong từng ô bảng (Table/Cell).
  - `🔄 Kiểm tra lại`: Kiểm tra nhanh lại đúng 1 lỗi đó bằng `Office.js` (không tốn lượt gọi AI).
  - `💬 Gắn Comment`: Chèn ghi chú Word trực tiếp vào vị trí lỗi.
  - `✏️ Sửa tự động (Track Changes)`: Bật `changeTrackingMode = trackAll` (nếu Word hỗ trợ) rồi thay thế văn bản.
  - `✖ Bỏ qua`: Mở hộp tùy chọn 2 cấp độ.
- **Cơ chế Bỏ qua 2 cấp độ:**
  1. `📄 Bỏ qua cho văn bản này`: Lưu tạm trong phiên làm việc (`Set` trên RAM, không gắn vào file Word). Hiển thị thanh **↺ Khôi phục** phía trên danh sách kết quả để hoàn tác nếu bấm nhầm.
  2. `🌐 Bỏ qua cho các văn bản khác`: Bắt buộc nhập lý do, lưu vào `localStorage`, vừa ẩn lỗi vừa gửi kèm lý do vào prompt AI ở các lần kiểm tra sau.
- **Quản lý danh sách bỏ qua chung:** Xem, xóa từng mục, **📥 Xuất file `.txt`**, và **📤 Nhập file `.txt`** theo cơ chế **gộp thêm + tự động lọc trùng** (không ghi đè danh sách đang có).

### 3.4. Công cụ Kiểm tra & Chuẩn hóa kỹ thuật theo NĐ 30 (2 bước có xác nhận)
- **Bước 1 — Quét 7 tiêu chí kỹ thuật (chưa tự sửa):**
  1. Khổ giấy A4 (`210 × 297 mm`).
  2. Căn lề 4 cạnh (Trên/Dưới `20–25 mm`, Trái `30–35 mm`, Phải `15–20 mm`).
  3. Phông chữ `Times New Roman` (quét cả trong bảng và ngoài bảng).
  4. Cỡ chữ nội dung (`13–14 pt`).
  5. Canh đều 2 lề (`Justify`).
  6. Thụt đầu dòng (`1 cm` hoặc `1,27 cm`).
  7. **Cách đoạn (`Before` & `After`) và Giãn dòng (`1,0–1,5 lines`):**
     - Báo cần chuẩn hóa nếu **một trong hai (`Before` hoặc `After`)** chưa khớp với mức chuẩn đang lưu (VD: `Before = 0 pt, After = 6 pt` vẫn báo lỗi để đồng bộ cả hai).
- **Bước 2 — Tích chọn, Chọn thông số đích & Xác nhận:**
  - Liệt kê các thành phần chưa đạt kèm **checkbox** (mặc định tích sẵn).
  - **8 ô chọn thông số đích theo đúng khung NĐ 30:**
    - Cỡ chữ nội dung: `13 pt` / `14 pt`.
    - Thụt đầu dòng: `1,27 cm` / `1,0 cm`.
    - Giãn dòng: `1,0` / `1,15` / `1,2` / `1,5 lines`.
    - **Cách đoạn (áp dụng giống nhau cho cả `Before` và `After`):** `3 pt` / `6 pt` / `8 pt` *(mặc định ban đầu là **`6 pt`**)*.
    - Lề trên & Lề dưới: `20` / `22` / `25 mm`; Lề trái: `30` / `32` / `35 mm`; Lề phải: `15` / `18` / `20 mm`.
  - **Nút `💾 Lưu làm mặc định`:** Lưu toàn bộ 8 thông số vào `localStorage`. Lần sau mở Word, Add-in tự động nạp sẵn 8 thông số này và dùng chính mức Cách đoạn đã lưu để đối chiếu đồng bộ `Before` và `After` ở bước Quét.
  - **Bảo vệ Bảng (Table) & Tiêu đề canh giữa:** Trong bảng (`parentTableOrNullObject`) **chỉ đổi phông chữ sang `Times New Roman`**, tuyệt đối không thay đổi cỡ chữ, thụt đầu dòng, canh đều hay khoảng cách đoạn.

---

## 4. YÊU CẦU KỸ THUẬT & TIÊU CHUẨN KIỂM THỬ BẮT BUỘC

1. **An toàn `Office.js` (tránh lỗi `The property 'items' is not available`):**
   - Trong hàm chuẩn hóa (`applySelectedFormatND30`), tách riêng khối `sections.load("items")` + `await context.sync()` và chỉ thực thi khi người dùng tích chọn `fmt_paper` hoặc `fmt_margin`.
   - Tách riêng `paragraphs.load("items")` + `await context.sync()`, sau đó mới load thuộc tính từng đoạn và `await context.sync()` trước khi đọc/ghi.
2. **Kiểm tra cú pháp tự động:** Chạy `node --check` trên toàn bộ file `.js` trước khi hoàn tất, đảm bảo 0 lỗi `SyntaxError`.
3. **Kịch bản kiểm thử (Acceptance Tests):**
   - **Test 1 (Offline 100%):** Ngắt mạng trước khi mở Word → Mở Add-in từ thẻ **Home** → Chạy **Kiểm tra 3 bước** → Bước 1 chạy đủ, Bước 2 tự chuyển sang chế độ Offline bắt đúng lỗi gõ sai & khẩu ngữ, Bước 3 bỏ qua êm không lỗi.
   - **Test 2 (Chuẩn hóa NĐ 30):** Chỉ tích chọn mục *Cách đoạn* (không tích Lề/Khổ giấy) → Bấm Xác nhận → Không báo lỗi `items`, cả `Before` và `After` đều được đặt bằng mức đã chọn, bố cục Bảng (Table) giữ nguyên.
   - **Test 3 (Lưu mặc định):** Đổi Cách đoạn sang `3 pt` + đổi 1 thông số lề → Bấm **💾 Lưu làm mặc định** → Reload Add-in → Xác nhận các ô chọn hiển thị đúng giá trị đã lưu và bước Quét đối chiếu đúng theo mức `3 pt`.