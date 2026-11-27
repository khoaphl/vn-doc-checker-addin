window.PCDT = window.PCDT || {};

(function (PCDT) {
  async function callGeminiJson(apiKey, prompt, stepLabel) {
    const models = ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-2.5-flash"];
    let lastError = "";

    for (const model of models) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

      for (let attempt = 1; attempt <= 2; attempt++) {
        PCDT.setStatus(`${stepLabel} (${model} - lần ${attempt})...`);
        const response = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
              responseMimeType: "application/json",
              temperature: 0.0
            }
          })
        });

        if (response.ok) {
          const data = await response.json();
          let rawText = data.candidates?.[0]?.content?.parts?.[0]?.text || "[]";
          rawText = rawText.replace(/^```json\s*/i, "").replace(/```$/i, "").trim();
          return JSON.parse(rawText);
        }

        lastError = `Lỗi Gemini API (${response.status}) với ${model}`;
        if (response.status === 503 || response.status === 429) {
          await new Promise((resolve) => setTimeout(resolve, 2000));
          continue;
        } else {
          break;
        }
      }
    }

    throw new Error(`${lastError}. Máy chủ Google đang bận, vui lòng thử lại sau ít phút.`);
  }

  async function detectDocumentType(apiKey, fullText) {
    const validTypes = [
      "Công văn",
      "Quyết định",
      "Tờ trình",
      "Báo cáo",
      "Thông báo",
      "Kế hoạch",
      "Biên bản",
      "Quy định / Quy chế",
      "Văn bản hành chính khác"
    ];

    const headerLines = fullText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .slice(0, 20);

    for (const line of headerLines) {
      if (/^QUYẾT ĐỊNH$/i.test(line)) return "Quyết định";
      if (/^TỜ TRÌNH$/i.test(line)) return "Tờ trình";
      if (/^BÁO CÁO$/i.test(line)) return "Báo cáo";
      if (/^THÔNG BÁO$/i.test(line)) return "Thông báo";
      if (/^KẾ HOẠCH$/i.test(line)) return "Kế hoạch";
      if (/^BIÊN BẢN$/i.test(line)) return "Biên bản";
      if (/^(QUY ĐỊNH|QUY CHẾ)$/i.test(line)) return "Quy định / Quy chế";
    }

    if (headerLines.some((l) => /^V\/v\b/i.test(l))) {
      return "Công văn";
    }

    if (!apiKey || !navigator.onLine) {
      return "Văn bản hành chính khác";
    }

    const detectPrompt = `
Xác định loại văn bản hành chính Việt Nam của đoạn trích dưới đây.
CHỈ chọn đúng 1 giá trị trong danh sách: ${JSON.stringify(validTypes)}.
Trả về JSON dạng: {"docType": "..."}

Đoạn trích văn bản:
"""
${fullText.slice(0, 1200)}
"""
`;
    try {
      const res = await callGeminiJson(apiKey, detectPrompt, "Đang tự động nhận diện loại văn bản");
      if (res && validTypes.includes(res.docType)) return res.docType;
    } catch (e) {}

    return "Văn bản hành chính khác";
  }

  // BƯỚC 2 (ONLINE): KIỂM TRA CHÍNH TẢ & NGÔN NGỮ
  async function runStep2Spelling(apiKey, documentText) {
    const globalIgnoreSection = PCDT.buildGlobalIgnorePromptSection();
    const prompt = `
Bạn là biên tập viên tiếng Việt chuyên rà soát văn bản hành chính.
NHIỆM VỤ DUY NHẤT: Chỉ kiểm tra lỗi CHÍNH TẢ, lỗi ĐÁNH MÁY, lỗi DẤU CÂU và từ ngữ KHẨU NGỮ/BIỆT NGỮ (như: vại, bít, thui, chúa nhật...).
KHÔNG kiểm tra thể thức, KHÔNG kiểm tra lề/font, KHÔNG kiểm tra pháp lý hay nơi nhận.
KHÔNG báo lỗi thiếu dấu cách giữa hai dòng khác nhau (vì văn bản đã được tách dòng theo bố cục trình bày).
Chỉ báo lỗi khi CHẮC CHẮN SAI (không bắt lỗi danh từ riêng, tên viết tắt ngành điện như EVN, EVNSPC, PCĐT, TBA, kV, MW...).
${globalIgnoreSection}
CHỈ TRẢ VỀ MỘT MẢNG JSON HỢP LỆ theo cấu trúc:
[
  {
    "category": "Chính tả",
    "quote": "trích dẫn chính xác 1 cụm từ ngắn có chứa lỗi trong văn bản (tối đa 60 ký tự, nằm trên cùng 1 dòng)",
    "issue": "mô tả ngắn gọn lỗi chính tả/đánh máy",
    "suggestion": "đề xuất cách sửa",
    "replacement": "chuỗi văn bản đúng để thay thế trực tiếp cho trường quote"
  }
]

Nếu không có lỗi chính tả nào, trả về [].

Nội dung văn bản:
"""
${documentText}
"""
`;
    const res = await callGeminiJson(apiKey, prompt, "Bước 2/3: Kiểm tra chính tả & ngôn ngữ");
    return Array.isArray(res) ? res : [];
  }

  // BƯỚC 3: KIỂM TRA CẤU TRÚC THỂ THỨC & TÍNH HỢP LỆ
  function getDocTypeSpecificChecklist(docType) {
    const checklists = {
      "Công văn": `
- Trích yếu (V/v): Phải phản ánh đầy đủ nội dung, viết chữ thường sau "V/v" (trừ danh từ riêng), không lồng tên loại ("V/v thông báo...", "V/v báo cáo...").
- Số, ký hiệu Công văn: KHÔNG có chữ "CV" trong ký hiệu (VD đúng: Số: 3156/PCĐT-VP).
- Kính gửi & Nơi nhận: Văn bản gửi ra ngoài Công ty không được ghi lẫn cả đơn vị bên ngoài và đơn vị nội bộ trong cùng mục "Kính gửi" (đơn vị nội bộ chuyển xuống "Nơi nhận"). Nơi nhận có "- Như trên;" nếu phù hợp.
- Không dùng Công văn để đặt quy tắc bắt buộc chung (quy chế/quy định) hoặc sửa nội dung văn bản khác (chỉ đính chính thể thức/kỹ thuật/căn cứ).
- Viện dẫn pháp lý: Lần đầu ghi đầy đủ (tên loại, số ký hiệu, ngày... tháng... năm..., cơ quan ban hành, trích yếu); lần sau chỉ ghi tên loại + số ký hiệu. Ngày tháng phải viết đầy đủ "ngày ... tháng ... năm ...".`,
      "Quyết định": `
- Phân biệt rõ 2 vị trí chức danh:
  + Vị trí 1 - THẨM QUYỀN BAN HÀNH (nằm ngay dưới trích yếu và TRÊN phần "Căn cứ...", ví dụ: "GIÁM ĐỐC CÔNG TY ĐIỆN LỰC ĐỒNG THÁP"): Đây KHÔNG phải chỗ ký. TUYỆT ĐỐI KHÔNG báo lỗi thiếu "KT."/"TM.", KHÔNG báo lỗi thiếu họ tên và KHÔNG yêu cầu chừa dòng ký tại vị trí này.
  + Vị trí 2 - NGƯỜI KÝ BAN HÀNH (chỉ nằm ở CUỐI văn bản, SAU ký hiệu "./." và ngang hàng mục "Nơi nhận:"): Chỉ kiểm tra quyền hạn ký (GIÁM ĐỐC hoặc KT. GIÁM ĐỐC / PHÓ GIÁM ĐỐC) và họ tên người ký tại vị trí cuối văn bản này. Nếu Giám đốc ký trực tiếp thì ghi "GIÁM ĐỐC" là ĐÚNG.
- Căn cứ pháp lý đầu tiên: Bắt buộc có Quyết định thành lập/đổi tên Công ty, Quy chế tổ chức và hoạt động, Quy chế phân cấp, sau đó mới đến các căn cứ liên quan khác. LƯU Ý QUAN TRỌNG: "Quyết định số 235/QĐ-EVN ngày 14 tháng 4 năm 2010 của Tập đoàn Điện lực Việt Nam về việc đổi tên các Điện lực trực thuộc Tổng công ty Điện lực miền Nam" CHÍNH LÀ Quyết định thành lập (đổi tên Công ty) hợp lệ — TUYỆT ĐỐI KHÔNG báo lỗi khi văn bản đặt Quyết định này ở căn cứ đầu tiên.
- Định dạng căn cứ: mỗi dòng kết thúc bằng ";", căn cứ cuối kết thúc bằng "." (hoặc ";" nếu tiếp theo là "Theo đề nghị của..." và dòng đó kết thúc bằng "."). Không dùng "Xét đề nghị của..." theo mẫu 1.2/1.3 NĐ 30.
- Điều khoản thi hành: Phải có Điều về hiệu lực thi hành và Điều về trách nhiệm thi hành; phần Nơi nhận có "- Như Điều ...;". Không thêm "Xin trân trọng cảm ơn./." ở cuối Quyết định.
- Viện dẫn Luật/Pháp lệnh: Không ghi số, ký hiệu, cơ quan ban hành.`,
      "Tờ trình": `
- Đúng đối tượng gửi: Chỉ dùng trình cấp trên có thẩm quyền phê duyệt, KHÔNG dùng Tờ trình gửi đơn vị cùng cấp.
- Mục nơi nhận đầu văn bản phải ghi "Kính gửi:" (không ghi "Kính trình:" ở đầu văn bản; "Kính trình ... xem xét, quyết định./." chỉ dùng ở câu kết). Không dùng "Xin trân trọng cảm ơn./.".
- Bố cục: Rõ lý do/sự cần thiết, nội dung đề xuất (phương án, kinh phí, tiến độ), kiến nghị phê duyệt cụ thể.`,
      "Báo cáo": `
- Bắt buộc có mục "Kính gửi:" gửi cơ quan/đơn vị trực tiếp chỉ đạo hoặc giải quyết.
- Bố cục: Kết quả đạt được → Tồn tại, hạn chế, nguyên nhân → Phương hướng, nhiệm vụ → Kiến nghị, đề xuất.
- Kiểm tra tính thống nhất của số liệu, thời gian kỳ báo cáo.`,
      "Thông báo": `
- Phải nêu rõ cơ sở, nội dung truyền đạt ý kiến chỉ đạo, kết luận (ai kết luận nếu là thông báo kết luận cuộc họp) hoặc sự việc cần phổ biến và đối tượng thực hiện.`,
      "Kế hoạch": `
- Bố cục rõ ràng: I. Mục đích, yêu cầu → II. Nội dung, nhiệm vụ (chủ trì, phối hợp, thời gian thực hiện) → III. Tổ chức thực hiện (kinh phí, trách nhiệm, chế độ báo cáo).`,
      "Biên bản": `
- Đầy đủ thông tin: Thời gian bắt đầu, địa điểm, thành phần tham dự (chủ trì, thư ký, có mặt/vắng mặt), diễn biến nội dung, kết luận, thời gian kết thúc.
- Khối ký: Thư ký (bên trái) và Chủ trì (bên phải).`,
      "Quy định / Quy chế": `
- Thường ban hành kèm theo Quyết định (ghi chỉ dẫn "(Kèm theo Quyết định số .../QĐ-... ngày ...)" nếu là bản giấy).
- Bố cục phân cấp chuẩn NĐ 30: Phần → Chương → Mục → Tiểu mục → Điều → Khoản → Điểm (a, b, c, d, đ, e, g, h...).`
    };
    return checklists[docType] || "- Kiểm tra đầy đủ Quốc hiệu, Tiêu ngữ, Tên cơ quan, Địa danh ngày tháng, Trích yếu, Viện dẫn pháp lý, Thẩm quyền ký và Nơi nhận.";
  }

  async function runStep3FormatAndValidity(apiKey, documentText, docType) {
    const customRules = localStorage.getItem("custom_rules_text") || "";
    const customRulesSection = customRules
      ? `\nQUY ĐỊNH NỘI BỘ BẮT BUỘC ƯU TIÊN ÁP DỤNG (nếu khác NĐ 30 thì ưu tiên theo quy định này):\n"""\n${customRules}\n"""\n`
      : "";
    const globalIgnoreSection = PCDT.buildGlobalIgnorePromptSection();
    const specificChecklist = getDocTypeSpecificChecklist(docType);

    const prompt = `
Bạn là chuyên gia văn thư lưu trữ và pháp chế tại Việt Nam (Nghị định 30/2020/NĐ-CP và quy định nội bộ ngành Điện).
LOẠI VĂN BẢN ĐANG KIỂM TRA: "${docType}".

NHIỆM VỤ: Chỉ kiểm tra CẤU TRÚC THỂ THỨC, VĂN PHONG CÔNG VỤ và TÍNH HỢP LỆ/LOGIC của văn bản theo đúng loại "${docType}".
TUYỆT ĐỐI KHÔNG lặp lại các lỗi đã kiểm tra ở bước trước: KHÔNG báo lỗi chính tả/đánh máy, KHÔNG báo lỗi font chữ/cỡ chữ/căn lề, KHÔNG báo lỗi thiếu "Số:" hay thiếu "./.".
KHÔNG tự bịa số liệu hoặc căn cứ thay thế khi đề xuất; nếu thiếu thông tin thì ghi rõ cần bổ sung.

Hãy rà soát tuần tự theo danh mục kiểm tra cố định sau:
1. Tên cơ quan ban hành, Địa danh và ngày tháng năm (viết đầy đủ ngày tháng, số < 10 và tháng 1, 2 thêm số 0).
2. Tiêu chí riêng cho loại "${docType}":
${specificChecklist}
3. Quyền hạn, chức vụ, họ tên người ký (không ghi học hàm/học vị trước tên).
4. Văn phong hành chính – công vụ (theo NĐ 30):
   - Tính minh xác: câu mệnh lệnh/giao việc phải rõ chủ thể thực hiện và thời hạn cụ thể (tránh viết chung chung kiểu "khẩn trương thực hiện" không rõ ai làm, khi nào xong).
   - Quan hệ gửi – nhận: phân biệt đúng "đề nghị" (cấp dưới/ngang cấp) và "yêu cầu/giao" (cấp trên).
   - Tránh từ ngữ biểu cảm, văn chương, sáo ngữ rườm rà lặp lại, dấu chấm than (!), dấu chấm lửng (...), emoji.
5. Tính hợp lệ & logic nội dung: mâu thuẫn ngày tháng, mâu thuẫn số liệu, viện dẫn văn bản không liên quan hoặc hết hiệu lực, sai thẩm quyền.
${customRulesSection}${globalIgnoreSection}
CHỈ TRẢ VỀ MỘT MẢNG JSON HỢP LỆ cho các mục CÓ LỖI theo cấu trúc:
[
  {
    "category": "Hợp lệ",
    "quote": "trích dẫn chính xác 1 cụm từ ngắn nằm trên cùng 1 dòng có trong văn bản để tìm kiếm (tối đa 80 ký tự), để rỗng nếu là lỗi thiếu thành phần",
    "issue": "mô tả ngắn gọn lỗi thể thức, văn phong hoặc tính hợp lệ",
    "suggestion": "đề xuất cách sửa cụ thể",
    "replacement": ""
  }
]

Nếu không phát hiện lỗi nào, trả về [].

Nội dung văn bản:
"""
${documentText}
"""
`;
    const res = await callGeminiJson(apiKey, prompt, `Bước 3/3: Kiểm tra thể thức & hợp lệ (${docType})`);
    return Array.isArray(res) ? res : [];
  }

  PCDT.callGeminiJson = callGeminiJson;
  PCDT.detectDocumentType = detectDocumentType;
  PCDT.runStep2Spelling = runStep2Spelling;
  PCDT.getDocTypeSpecificChecklist = getDocTypeSpecificChecklist;
  PCDT.runStep3FormatAndValidity = runStep3FormatAndValidity;
})(window.PCDT);
