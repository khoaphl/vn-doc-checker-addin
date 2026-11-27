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

      // Bổ sung từ QĐ 3738/QĐ-PCĐT, 981/QĐ-PCĐT, 3219/QĐ-PCĐT: Kiểm tra Cơ quan chủ quản - Đơn vị ban hành - Số hiệu (Cấp Công ty vs Phòng/Ban & Đơn vị trực thuộc)
      const pcdtDepts = [
        { pattern: /Phòng\s+Tổ\s+chức\s+và\s+Nhân\s+sự/i, name: "Phòng Tổ chức và Nhân sự", code: "TCNS" },
        { pattern: /Phòng\s+Kế\s+hoạch\s+và\s+Vật\s+tư/i, name: "Phòng Kế hoạch và Vật tư", code: "KHVT" },
        { pattern: /Phòng\s+Tài\s+chính\s+Kế\s+toán/i, name: "Phòng Tài chính Kế toán", code: "TCKT" },
        { pattern: /Phòng\s+Quản\s+lý\s+đầu\s+tư/i, name: "Phòng Quản lý đầu tư", code: "ĐT" },
        { pattern: /Phòng\s+Kỹ\s+thuật/i, name: "Phòng Kỹ thuật", code: "KT" },
        { pattern: /Phòng\s+Kinh\s+doanh/i, name: "Phòng Kinh doanh", code: "KD" },
        { pattern: /Phòng\s+An\s+toàn/i, name: "Phòng An toàn", code: "AT" },
        { pattern: /Phòng\s+Điều\s+độ/i, name: "Phòng Điều độ", code: "ĐĐ" },
        { pattern: /Phòng\s+Kiểm\s+tra,?\s+thanh\s+tra\s+và\s+Pháp\s+chế/i, name: "Phòng Kiểm tra, thanh tra và Pháp chế", code: "KTTT" },
        { pattern: /Phòng\s+Khoa\s+học\s+công\s+nghệ\s+và\s+Chuyển\s+đổi\s+số/i, name: "Phòng Khoa học công nghệ và Chuyển đổi số", code: "VTCNTT" },
        { pattern: /Phòng\s+Viễn\s+thông\s+và\s+Công\s+nghệ\s+thông\s+tin/i, name: "Phòng Viễn thông và CNTT", code: "VTCNTT" },
        { pattern: /Ban\s+Quản\s+lý\s+dự\s+án/i, name: "Ban Quản lý dự án", code: "QLDA" },
        { pattern: /^Văn\s+phòng$/i, name: "Văn phòng", code: "VP", lineMatchOnly: true }
      ];

      const pcdtUnits = [
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Cao\s+Lãnh/i, name: "Đội QLĐ/Điện lực Cao Lãnh", codes: ["ĐCL"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Mỹ\s+Thọ/i, name: "Đội QLĐ/Điện lực Mỹ Thọ", codes: ["ĐMT", "ĐMTh"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Tháp\s+Mười/i, name: "Đội QLĐ/Điện lực Tháp Mười", codes: ["ĐTM"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Thanh\s+Bình/i, name: "Đội QLĐ/Điện lực Thanh Bình", codes: ["ĐTB"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Tam\s+Nông/i, name: "Đội QLĐ/Điện lực Tam Nông", codes: ["ĐTN"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Hồng\s+Ngự/i, name: "Đội QLĐ/Điện lực Hồng Ngự", codes: ["ĐHN"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Tân\s+Hồng/i, name: "Đội QLĐ/Điện lực Tân Hồng", codes: ["ĐTH"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Thường\s+Thới\s+Tiền/i, name: "Đội QLĐ/Điện lực Thường Thới Tiền", codes: ["ĐTTT"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Sa\s+Đéc/i, name: "Đội QLĐ/Điện lực Sa Đéc", codes: ["ĐSĐ"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Nha\s+Mân/i, name: "Đội QLĐ/Điện lực Nha Mân", codes: ["ĐNM"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Lai\s+Vung/i, name: "Đội QLĐ/Điện lực Lai Vung", codes: ["ĐLV"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Lấp\s+Vò/i, name: "Đội QLĐ/Điện lực Lấp Vò", codes: ["ĐLVO"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Tân\s+Phú\s+Đông/i, name: "Đội QLĐ/Điện lực Tân Phú Đông", codes: ["ĐTPĐ"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Cái\s+Bè/i, name: "Đội QLĐ/Điện lực Cái Bè", codes: ["ĐCB"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Cai\s+Lậy/i, name: "Đội QLĐ/Điện lực Cai Lậy", codes: ["ĐCLa"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Tân\s+Phước/i, name: "Đội QLĐ/Điện lực Tân Phước", codes: ["ĐTP"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Châu\s+Thành/i, name: "Đội QLĐ/Điện lực Châu Thành", codes: ["ĐCT"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Chợ\s+Gạo/i, name: "Đội QLĐ/Điện lực Chợ Gạo", codes: ["ĐCG"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Vĩnh\s+Bình/i, name: "Đội QLĐ/Điện lực Vĩnh Bình", codes: ["ĐVB"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Gò\s+Công\s+Đông/i, name: "Đội QLĐ/Điện lực Gò Công Đông", codes: ["ĐGCĐ"] },
        { pattern: /(?:Điện\s+lực|Đội\s+Quản\s+lý\s+điện)\s+Gò\s+Công/i, name: "Đội QLĐ/Điện lực Gò Công", codes: ["ĐGC"] },
        { pattern: /Xí\s+nghiệp\s+Lưới\s+điện\s+cao\s+thế\s+Đồng\s+Tháp/i, name: "Xí nghiệp Lưới điện cao thế Đồng Tháp", codes: ["XNCT"] },
        { pattern: /Xí\s+nghiệp\s+Dịch\s+vụ\s+Điện\s+lực\s+Đồng\s+Tháp/i, name: "Xí nghiệp Dịch vụ Điện lực Đồng Tháp", codes: ["XNDV"] },
        { pattern: /Điện\s+lực\s+Bình\s+Phú/i, name: "Điện lực Bình Phú", codes: ["ĐLBP", "BPH"] }
      ];

      const allLines = fullText.split(/\r?\n/);
      const headerEndIdx = allLines.findIndex((l) =>
        /^(Số\s*:|V\/v\b|Kính\s+gửi\b|Căn\s+cứ\b|QUYẾT\s+ĐỊNH\b|TỜ\s+TRÌNH\b|BÁO\s+CÁO\b|THÔNG\s+BÁO\b|KẾ\s+HOẠCH\b|BIÊN\s+BẢN\b)/i.test(l.trim())
      );
      const headerLines = allLines.slice(0, headerEndIdx > 0 ? Math.min(headerEndIdx, 14) : 10);
      const headerFlat = headerLines.join(" ").replace(/\s+/g, " ").trim();
      const hasTct = /TỔNG\s+CÔNG\s+TY\s+ĐIỆN\s+LỰC\s+MIỀN\s+NAM/i.test(headerFlat);
      const hasPcdt = /CÔNG\s+TY\s+ĐIỆN\s+LỰC\s+ĐỒNG\s+THÁP/i.test(headerFlat);

      const matchedDept = pcdtDepts.find((d) =>
        d.lineMatchOnly ? headerLines.some((l) => d.pattern.test(l.trim())) : d.pattern.test(headerFlat)
      );
      const matchedUnit = !matchedDept ? pcdtUnits.find((u) => u.pattern.test(headerFlat)) : null;

      const soLine = allLines.slice(0, 20).find((l) => /Số\s*:/i.test(l)) || "";
      const soCodeMatch = soLine.match(/Số\s*:\s*[^\r\n\/]*\/([A-ZĐÂÊÔƠƯĂa-zđ0-9+-]+)/i);
      const soSuffix = soCodeMatch ? soCodeMatch[1].trim() : "";

      if (matchedDept) {
        if (!hasPcdt || hasTct) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: (headerLines.find((l) => matchedDept.pattern.test(l)) || headerLines[0] || "").slice(0, 70),
            issue: `Văn bản cấp Phòng (${matchedDept.name}) phải ghi cơ quan cấp trên là 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP' (không ghi 'TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM').`,
            suggestion: `Đổi phần cơ quan chủ quản phía trên thành 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP' và đơn vị ban hành là '${matchedDept.name.toUpperCase()}'.`,
            replacement: ""
          });
        }
        if (soSuffix) {
          if (/^PCĐT-/i.test(soSuffix)) {
            issues.push({
              category: "Thể thức",
              checkType: "text",
              quote: soLine.trim().slice(0, 70),
              issue: `Văn bản cấp Phòng (${matchedDept.name}) không dùng ký hiệu '/PCĐT-...' của cấp Công ty.`,
              suggestion: `Đổi số, ký hiệu sang cấp Phòng: Số: .../${matchedDept.code} (hoặc .../<Loại VB>-${matchedDept.code} nếu là văn bản có tên loại).`,
              replacement: ""
            });
          } else {
            const deptCodesPattern = new RegExp(`(^|-)${matchedDept.code}$`, "i");
            if (!deptCodesPattern.test(soSuffix)) {
              issues.push({
                category: "Thể thức",
                checkType: "text",
                quote: soLine.trim().slice(0, 70),
                issue: `Ký hiệu văn bản '${soSuffix}' chưa khớp với '${matchedDept.name}' (ký hiệu chuẩn là '${matchedDept.code}').`,
                suggestion: `Sửa ký hiệu thành Số: .../${matchedDept.code}.`,
                replacement: ""
              });
            }
          }
        }
      } else if (matchedUnit) {
        if (!hasPcdt || hasTct) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: (headerLines.find((l) => matchedUnit.pattern.test(l)) || headerLines[0] || "").slice(0, 70),
            issue: `Văn bản của đơn vị trực thuộc (${matchedUnit.name}) phải ghi cơ quan cấp trên là 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP'.`,
            suggestion: `Đổi cơ quan chủ quản phía trên thành 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP'.`,
            replacement: ""
          });
        }
        if (soSuffix && /^PCĐT-/i.test(soSuffix)) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: soLine.trim().slice(0, 70),
            issue: `Văn bản của ${matchedUnit.name} không dùng ký hiệu '/PCĐT-...' của cấp Công ty.`,
            suggestion: `Dùng ký hiệu của đơn vị (${matchedUnit.codes.join(" hoặc ")}) theo QĐ 3738/QĐ-PCĐT & QĐ 3219/QĐ-PCĐT.`,
            replacement: ""
          });
        }
      } else if (hasPcdt) {
        if (!hasTct) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: (headerLines.find((l) => /CÔNG\s+TY\s+ĐIỆN\s+LỰC\s+ĐỒNG\s+THÁP/i.test(l)) || "").slice(0, 70),
            issue: "Văn bản cấp Công ty ('CÔNG TY ĐIỆN LỰC ĐỒNG THÁP') thiếu cơ quan cấp trên 'TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM'.",
            suggestion: "Bổ sung dòng 'TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM' phía trên 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP'.",
            replacement: ""
          });
        }
        if (docType === "Công văn" && soSuffix && !/^PCĐT-/i.test(soSuffix)) {
          const isDeptOnlyCode = pcdtDepts.some((d) => d.code.toUpperCase() === soSuffix.toUpperCase());
          if (isDeptOnlyCode) {
            issues.push({
              category: "Thể thức",
              checkType: "text",
              quote: soLine.trim().slice(0, 70),
              issue: `Công văn cấp Công ty đang dùng ký hiệu cấp Phòng '/${soSuffix}' (thiếu 'PCĐT-').`,
              suggestion: `Nếu do Công ty ban hành thì ghi 'Số: .../PCĐT-${soSuffix}'; nếu là văn bản nội bộ cấp Phòng thì đổi đơn vị ban hành thành tên Phòng và cấp trên là 'CÔNG TY ĐIỆN LỰC ĐỒNG THÁP'.`,
              replacement: ""
            });
          }
        }
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

        // Kiểm tra phần viện dẫn pháp lý theo Công văn 9975/PCĐT-VP (Mục II.2.b)
        const dieu1Boundary = fullText.search(/^[ \t]*Điều\s+1\b/im);
        const preambleText = dieu1Boundary > 0 ? fullText.slice(0, dieu1Boundary) : fullText.slice(0, 5000);
        const firstCanCuMatch = preambleText.match(/^[ \t]*Căn cứ[^\r\n]{0,65}/im);
        const fallbackQuote = firstCanCuMatch
          ? firstCanCuMatch[0].trim().slice(0, 70)
          : (preambleText.match(/^[ \t]*QUYẾT ĐỊNH[^\r\n]{0,55}/im)?.[0]?.trim().slice(0, 70) || "");

        const evnSpcSpaceMatch = preambleText.match(/[^\r\n]{0,25}EVN\s+SPC[^\r\n]{0,25}/i);
        if (evnSpcSpaceMatch) {
          const badSpc = evnSpcSpaceMatch[0].trim();
          issues.push({
            category: "Kỹ thuật",
            checkType: "text",
            quote: badSpc.slice(0, 70),
            issue: "Viết thừa khoảng trắng trong ký hiệu 'EVN SPC' ở phần căn cứ.",
            suggestion: "Viết liền 'EVNSPC' (ví dụ: 1677/QĐ-EVNSPC, 1824/QĐ-EVNSPC).",
            replacement: badSpc.replace(/EVN\s+SPC/gi, "EVNSPC")
          });
        }

        const foundationalBases = [
          {
            key: "235/QĐ-EVN",
            regex: /235\s*\/\s*Q[ĐÐ]-EVN\b/i,
            label: "Căn cứ 1: Quyết định số 235/QĐ-EVN ngày 14 tháng 4 năm 2010 của EVN"
          },
          {
            key: "201/QĐ-HĐTV",
            regex: /201\s*\/\s*Q[ĐÐ]-HĐTV\b/i,
            label: "Căn cứ 2: Quyết định số 201/QĐ-HĐTV ngày 05 tháng 8 năm 2025 của HĐTV EVNSPC"
          },
          {
            key: "1677/QĐ-EVNSPC",
            regex: /1677\s*\/\s*Q[ĐÐ]-EVN\s*SPC\b/i,
            label: "Căn cứ 3: Quyết định số 1677/QĐ-EVNSPC ngày 01 tháng 6 năm 2025 của EVNSPC"
          },
          {
            key: "1824/QĐ-EVNSPC",
            regex: /1824\s*\/\s*Q[ĐÐ]-EVN\s*SPC\b/i,
            label: "Căn cứ 3 (cùng dòng): Quyết định số 1824/QĐ-EVNSPC ngày 17 tháng 6 năm 2025 của EVNSPC"
          }
        ];

        const standardThreeCanCu = [
          "1) Căn cứ Quyết định số 235/QĐ-EVN ngày 14 tháng 4 năm 2010 của Tập đoàn Điện lực Việt Nam về việc đổi tên các Điện lực trực thuộc Tổng công ty Điện lực miền Nam;",
          "2) Căn cứ Quyết định số 201/QĐ-HĐTV ngày 05 tháng 8 năm 2025 của Hội đồng thành viên Tổng công ty Điện lực miền Nam về việc ban hành Quy chế tổ chức và hoạt động của Công ty Điện lực Đồng Tháp;",
          "3) Căn cứ Quyết định số 1677/QĐ-EVNSPC ngày 01 tháng 6 năm 2025 của Tổng công ty Điện lực miền Nam về việc ban hành Quy chế phân cấp quản lý giữa Tổng Giám đốc và Giám đốc các đơn vị trực thuộc trong Tổng công ty Điện lực miền Nam và Quyết định số 1824/QĐ-EVNSPC ngày 17 tháng 6 năm 2025 của Tổng công ty Điện lực miền Nam về việc sửa đổi, bổ sung một số nội dung trong “Quy chế phân cấp thẩm quyền quản lý giữa Tổng Giám đốc và Giám đốc đơn vị trực thuộc trong Tổng công ty Điện lực miền Nam” được ban hành theo Quyết định số 1677/QĐ-EVNSPC ngày 01/06/2025;"
        ].join(" ");

        const foundPositions = foundationalBases.map((b) => {
          const m = preambleText.match(b.regex);
          return { ...b, index: m ? m.index : -1 };
        });

        const missingBases = foundPositions.filter((b) => b.index === -1);
        if (missingBases.length > 0) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: fallbackQuote,
            issue: `Phần viện dẫn Quyết định thiếu căn cứ bắt buộc theo Công văn 9975/PCĐT-VP: ${missingBases.map((b) => b.key).join(", ")}.`,
            suggestion: `Bổ sung đầy đủ 3 căn cứ bắt buộc lên đầu phần viện dẫn theo đúng mẫu quy định: ${standardThreeCanCu}`,
            replacement: ""
          });
        }

        // Kiểm tra thứ tự: 235/QĐ-EVN phải đứng ở căn cứ đầu tiên, sau đó lần lượt 201/QĐ-HĐTV -> 1677/QĐ-EVNSPC -> 1824/QĐ-EVNSPC
        if (firstCanCuMatch && foundPositions[0].index !== -1) {
          const firstCanCuIdx = firstCanCuMatch.index;
          const nextCanCuRelIdx = preambleText
            .slice(firstCanCuIdx + firstCanCuMatch[0].length)
            .search(/^[ \t]*(Căn cứ|Theo đề nghị|Xét đề nghị|Xét\b)(?=\s|$)/im);
          const firstBlockEnd = nextCanCuRelIdx !== -1
            ? firstCanCuIdx + firstCanCuMatch[0].length + nextCanCuRelIdx
            : preambleText.length;

          if (foundPositions[0].index > firstBlockEnd) {
            issues.push({
              category: "Thể thức",
              checkType: "text",
              quote: firstCanCuMatch[0].trim().slice(0, 70),
              issue: "Thứ tự viện dẫn chưa đúng Công văn 9975/PCĐT-VP: Đặt căn cứ pháp lý khác trước Quyết định thành lập (235/QĐ-EVN).",
              suggestion: "Đưa các căn cứ nền tảng lên đầu phần viện dẫn theo thứ tự: (1) QĐ 235/QĐ-EVN -> (2) QĐ 201/QĐ-HĐTV -> (3) QĐ 1677/QĐ-EVNSPC & 1824/QĐ-EVNSPC -> các căn cứ khác.",
              replacement: ""
            });
          }
        }

        const presentBases = foundPositions.filter((b) => b.index !== -1);
        let isOrderWrong = false;
        for (let i = 1; i < presentBases.length; i++) {
          if (presentBases[i].index < presentBases[i - 1].index) {
            isOrderWrong = true;
            break;
          }
        }
        if (isOrderWrong && fallbackQuote) {
          issues.push({
            category: "Thể thức",
            checkType: "text",
            quote: fallbackQuote,
            issue: "Thứ tự các căn cứ nền tảng trong phần viện dẫn chưa đúng theo Công văn 9975/PCĐT-VP.",
            suggestion: "Sắp xếp lại đúng thứ tự: (1) QĐ 235/QĐ-EVN (Thành lập) -> (2) QĐ 201/QĐ-HĐTV (Quy chế TC&HĐ) -> (3) QĐ 1677/QĐ-EVNSPC & 1824/QĐ-EVNSPC (Quy chế phân cấp).",
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
