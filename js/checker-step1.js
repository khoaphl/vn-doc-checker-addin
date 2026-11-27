window.PCDT = window.PCDT || {};

(function (PCDT) {
  // BƯỚC 1: KIỂM TRA KỸ THUẬT & THỂ THỨC CỨNG (100% CỤC BỘ, KHÔNG GỌI AI)
  async function runStep1TechnicalAndHardRules(docType) {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      const sections = context.document.sections;

      paragraphs.load("items/text, items/font/name, items/font/size");

      const supportsPageSetup = Office.context.requirements.isSetSupported("WordApi", "1.5");
      if (supportsPageSetup) {
        sections.load("items/pageSetup/topMargin, items/pageSetup/bottomMargin, items/pageSetup/leftMargin, items/pageSetup/rightMargin");
      }

      await context.sync();

      const fullText = paragraphs.items
        .map((p) => (p.text || "").replace(/[\v\u000b\r]/g, "\n").trim())
        .filter((l) => l.length > 0)
        .join("\n");

      const issues = [];

      // 1. Căn lề trang (NĐ 30)
      if (supportsPageSetup) {
        sections.items.forEach((sec, idx) => {
          const ps = sec.pageSetup;
          const top = PCDT.ptToMm(ps.topMargin);
          const bottom = PCDT.ptToMm(ps.bottomMargin);
          const left = PCDT.ptToMm(ps.leftMargin);
          const right = PCDT.ptToMm(ps.rightMargin);

          const marginErrors = [];
          if (top < 20 || top > 25) marginErrors.push(`Lề trên ${top}mm (chuẩn 20–25mm)`);
          if (bottom < 20 || bottom > 25) marginErrors.push(`Lề dưới ${bottom}mm (chuẩn 20–25mm)`);
          if (left < 30 || left > 35) marginErrors.push(`Lề trái ${left}mm (chuẩn 30–35mm)`);
          if (right < 15 || right > 20) marginErrors.push(`Lề phải ${right}mm (chuẩn 15–20mm)`);

          if (marginErrors.length > 0) {
            issues.push({
              category: "Kỹ thuật",
              checkType: "margin",
              sectionIndex: idx,
              quote: "",
              issue: `Phần (Section) ${idx + 1} sai căn lề NĐ 30: ${marginErrors.join("; ")}.`,
              suggestion: "Bấm nút 'Kiểm tra & Chuẩn hóa (NĐ 30)' hoặc chỉnh thủ công (Trên/Dưới 20–25mm, Trái 30–35mm, Phải 15–20mm).",
              replacement: ""
            });
          }
        });
      }

      // 2. Phông chữ, cỡ chữ & ký tự Markdown/Chatbot sót lại
      paragraphs.items.forEach((p, idx) => {
        const text = (p.text || "").replace(/[\v\u000b\r]/g, " ").trim();
        if (!text) return;

        if (p.font.name && p.font.name !== "Times New Roman") {
          issues.push({
            category: "Kỹ thuật",
            checkType: "font",
            quote: text.slice(0, 60),
            issue: `Đoạn ${idx + 1} đang dùng font "${p.font.name}" thay vì "Times New Roman".`,
            suggestion: "Chuyển sang font Times New Roman (hoặc dùng nút Kiểm tra & Chuẩn hóa NĐ 30).",
            replacement: ""
          });
        }

        if (p.font.size && (p.font.size < 12 || p.font.size > 14) && text.length > 40) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "fontSize",
            quote: text.slice(0, 60),
            issue: `Đoạn nội dung ${idx + 1} có cỡ chữ ${p.font.size}pt (chuẩn nội dung là 13pt hoặc 14pt).`,
            suggestion: "Chỉnh cỡ chữ về 13pt hoặc 14pt.",
            replacement: ""
          });
        }

        if (/^\s*•\s+/.test(text)) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: text.slice(0, 60),
            issue: "Dùng dấu chấm tròn (•) kiểu liệt kê chatbot/web thay cho ký hiệu liệt kê NĐ 30.",
            suggestion: "Thay bằng gạch đầu dòng (-), khoản (1., 2.) hoặc điểm (a), b)) theo NĐ 30.",
            replacement: ""
          });
        }

        if (/^-{3,}$/.test(text)) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: text,
            issue: "Còn sót đường kẻ ngang Markdown (---) trong văn bản Word.",
            suggestion: "Xóa ký tự '---' và dùng đường kẻ chuẩn của Word nếu cần phân cách.",
            replacement: ""
          });
        }
      });

      // 3. Quy tắc cứng D-Office EVN & PCĐT
      if (!/Số\s*:/i.test(fullText)) {
        issues.push({
          category: "Kỹ thuật",
          checkType: "missingSo",
          quote: "",
          issue: "Thiếu cụm từ 'Số:' để hệ thống D-Office EVN nhận dạng và cấp số tự động.",
          suggestion: "Bổ sung dòng 'Số: .../...' đúng vị trí thể thức.",
          replacement: ""
        });
      }

      if (!fullText.includes("./.")) {
        issues.push({
          category: "Kỹ thuật",
          checkType: "missingEnd",
          quote: "",
          issue: "Thiếu ký hiệu kết thúc nội dung './.' (bắt buộc để hệ thống D-Office chèn chữ ký tắt).",
          suggestion: "Thêm ký hiệu './.' vào cuối dòng kết thúc phần nội dung văn bản.",
          replacement: ""
        });
      }

      // Bắt lỗi viết tắt viện dẫn văn bản (CV1041, TB134, VB294...)
      const badCitationMatches = fullText.match(/\b(CV|TB|VB)\s*\d+\b/g);
      if (badCitationMatches) {
        const uniqueMatches = [...new Set(badCitationMatches)];
        uniqueMatches.forEach((m) => {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: m,
            issue: `Viết tắt viện dẫn văn bản "${m}" sai quy định (CV 9975/PCĐT-VP).`,
            suggestion: "Ghi đầy đủ tên loại và số, ký hiệu (ví dụ: Công văn số 1041/EVNSPC-VTCNTT).",
            replacement: ""
          });
        });
      }

      // Bổ sung từ NĐ 30 (vietnamese-docs-style): Công văn không có chữ "CV" trong số/ký hiệu
      if (docType === "Công văn") {
        const badCvCode = fullText.match(/Số\s*:\s*[^\r\n]*\/CV-/i);
        if (badCvCode) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: badCvCode[0].slice(0, 60),
            issue: "Số, ký hiệu Công văn có chứa '/CV-' (sai theo Phụ lục I NĐ 30).",
            suggestion: "Công văn không ghi chữ 'CV' trong ký hiệu (VD đúng: Số: 3156/PCĐT-VP).",
            replacement: ""
          });
        }

        const badSubject = fullText.match(/V\/v\s+(thông báo|báo cáo)[^\r\n]{0,50}/i);
        if (badSubject) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: badSubject[0].slice(0, 80),
            issue: "Trích yếu Công văn có chứa từ 'thông báo' hoặc 'báo cáo' (sai theo CV 9975/PCĐT-VP).",
            suggestion: "Bỏ từ 'thông báo'/'báo cáo' trong trích yếu Công văn, hoặc chuyển sang mẫu văn bản có tên loại THÔNG BÁO / BÁO CÁO.",
            replacement: ""
          });
        }
      }

      // Bổ sung từ NĐ 30: Trong Quyết định, mẫu 1.2/1.3 dùng "Theo đề nghị của..." chứ không dùng "Xét đề nghị của..."
      if (docType === "Quyết định") {
        const xetDeNghi = fullText.match(/^Xét đề nghị của[^\r\n]{0,60}/im);
        if (xetDeNghi) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: xetDeNghi[0].slice(0, 70),
            issue: "Dùng 'Xét đề nghị của...' trong phần căn cứ Quyết định (theo mẫu 1.2/1.3 NĐ 30 chuẩn là 'Theo đề nghị của...').",
            suggestion: "Đổi thành 'Theo đề nghị của ...' và kết thúc bằng dấu chấm (.).",
            replacement: ""
          });
        }
      }

      // Bổ sung từ NĐ 30: Kiểm tra thiếu số 0 ở dòng địa danh ngày tháng (ngày < 10, tháng 1 hoặc 2) trong 25 dòng đầu
      const headLines = fullText.split(/\r?\n/).slice(0, 25);
      headLines.forEach((line) => {
        const mDate = line.match(/ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+\d{4}/i);
        if (mDate) {
          const dayStr = mDate[1];
          const monthStr = mDate[2];
          if ((dayStr.length === 1 && parseInt(dayStr, 10) < 10) || (monthStr.length === 1 && (monthStr === "1" || monthStr === "2"))) {
            issues.push({
              category: "Kỹ thuật",
              checkType: "text",
              quote: mDate[0],
              issue: `Ngày/tháng ban hành "${mDate[0]}" thiếu số 0 ở phía trước (NĐ 30: ngày < 10 và tháng 1, 2 phải thêm số 0).`,
              suggestion: `Viết đủ 2 chữ số: ngày ${dayStr.padStart(2, "0")} tháng ${monthStr.padStart(2, "0")}.`,
              replacement: ""
            });
          }
        }
      });

      // 4. 15 dòng cuối: Chức vụ người ký bị viết tắt hoặc sai tên gọi
      const tailLines = fullText
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean)
        .slice(-15);
      tailLines.forEach((line) => {
        if (/^(KT\.\s*)?PGĐ\b/.test(line) || /\bPHÓ ĐIỆN LỰC\b/i.test(line)) {
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: line.slice(0, 80),
            issue: `Chức vụ người ký "${line}" bị viết tắt hoặc sai tên gọi (CV 3156/PCĐT-VP).`,
            suggestion: "Viết đầy đủ chức vụ in hoa: PHÓ GIÁM ĐỐC hoặc PHÓ TRƯỞNG ĐIỆN LỰC.",
            replacement: ""
          });
        }
      });

      return { fullText, step1Issues: issues };
    });
  }

  // KIỂM TRA NHANH LẠI 1 LỖI CỤ THỂ (KHÔNG TỐN LƯỢT GỌI AI)
  async function recheckSingleIssue(item) {
    return await Word.run(async (context) => {
      if (item.checkType === "margin") {
        const supportsPageSetup = Office.context.requirements.isSetSupported("WordApi", "1.5");
        if (!supportsPageSetup) return { fixed: false, message: "Phiên bản Word không hỗ trợ đọc lề trang." };

        const sections = context.document.sections;
        sections.load("items/pageSetup/topMargin, items/pageSetup/bottomMargin, items/pageSetup/leftMargin, items/pageSetup/rightMargin");
        await context.sync();

        const sec = sections.items[item.sectionIndex || 0];
        if (!sec) return { fixed: true, message: "Đã đạt chuẩn căn lề." };

        const top = PCDT.ptToMm(sec.pageSetup.topMargin);
        const bottom = PCDT.ptToMm(sec.pageSetup.bottomMargin);
        const left = PCDT.ptToMm(sec.pageSetup.leftMargin);
        const right = PCDT.ptToMm(sec.pageSetup.rightMargin);

        const ok = top >= 20 && top <= 25 && bottom >= 20 && bottom <= 25 && left >= 30 && left <= 35 && right >= 15 && right <= 20;
        return ok
          ? { fixed: true, message: "Lề trang đã đúng chuẩn NĐ 30." }
          : { fixed: false, message: `Lề vẫn chưa đúng (Hiện tại: Trên ${top}, Dưới ${bottom}, Trái ${left}, Phải ${right}mm).` };
      }

      if (item.checkType === "font") {
        const target = await PCDT.findTargetInWord(context, item.quote);
        if (!target) return { fixed: true, message: "Đoạn văn bản lỗi không còn tồn tại." };
        const range = typeof target.getRange === "function" ? target.getRange() : target;
        range.font.load("name");
        await context.sync();
        return range.font.name === "Times New Roman"
          ? { fixed: true, message: "Font chữ đã chuyển về Times New Roman." }
          : { fixed: false, message: `Vẫn đang dùng font "${range.font.name}".` };
      }

      if (item.checkType === "fontSize") {
        const target = await PCDT.findTargetInWord(context, item.quote);
        if (!target) return { fixed: true, message: "Đoạn văn bản lỗi không còn tồn tại." };
        const range = typeof target.getRange === "function" ? target.getRange() : target;
        range.font.load("size");
        await context.sync();
        return range.font.size >= 12 && range.font.size <= 14
          ? { fixed: true, message: `Cỡ chữ đã đạt chuẩn (${range.font.size}pt).` }
          : { fixed: false, message: `Cỡ chữ hiện vẫn là ${range.font.size}pt.` };
      }

      if (item.checkType === "missingSo") {
        const cleanText = await PCDT.extractCleanDocumentText(context);
        return /Số\s*:/i.test(cleanText)
          ? { fixed: true, message: "Đã bổ sung cụm từ 'Số:'." }
          : { fixed: false, message: "Vẫn chưa tìm thấy cụm từ 'Số:' trong văn bản." };
      }

      if (item.checkType === "missingEnd") {
        const cleanText = await PCDT.extractCleanDocumentText(context);
        return cleanText.includes("./.")
          ? { fixed: true, message: "Đã bổ sung ký hiệu kết thúc './.'." }
          : { fixed: false, message: "Vẫn chưa thấy ký hiệu './.' trong văn bản." };
      }

      if (item.quote) {
        const cleanText = await PCDT.extractCleanDocumentText(context);
        const bodyNorm = PCDT.normalizeText(cleanText);
        const quoteParts = item.quote
          .split(/[\r\n\t\v\f\u000b]+/)
          .map(PCDT.normalizeText)
          .filter((s) => s.length >= 2);

        const stillExists = quoteParts.length > 0 && quoteParts.every((part) => bodyNorm.includes(part));
        return !stillExists
          ? { fixed: true, message: "Đoạn văn bản lỗi cũ đã được chỉnh sửa." }
          : { fixed: false, message: "Đoạn văn bản lỗi cũ vẫn còn trong tài liệu." };
      }

      return { fixed: false, message: "Mục này không có trích dẫn cụ thể để kiểm tra tự động." };
    });
  }

  PCDT.runStep1TechnicalAndHardRules = runStep1TechnicalAndHardRules;
  PCDT.recheckSingleIssue = recheckSingleIssue;
})(window.PCDT);
